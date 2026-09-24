using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using Doctracker.Core.Services;
using System.Windows.Forms;
using Doctracker.AddIn.Excel;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed partial class DoctrackerPaneControl : UserControl
    {
        private readonly ExcelInterop.Application application;
        private readonly WorkbookProjectContext context;
        private readonly ExcelCellGateway cells;
        private readonly IOcrEngine ocr;
        private readonly WorkspaceView view;
        private bool bindingProofs;
        private string focusedSnipId;
        private string lastImportedId;
        private readonly DocumentCanvas canvas;
        private readonly ComboBox documents;
        private TextBox searchBox;
        private readonly ListBox searchResults;
        private readonly Label ocrState;
        private Label modeState;
        private readonly Label status;

        private SnipType? activeSnipType;
        private readonly ExcelInterop.Workbook workbook;
        private CancellationTokenSource operation;
        private string boundProjectPath;
        private readonly Button cancelButton;
        private string matchingInputSheet;
        private string matchingInputAddress;
        private string matchingOutputSheet;
        private string matchingOutputAddress;

        public DoctrackerPaneControl(ExcelInterop.Application application, ExcelInterop.Workbook workbook, WorkbookProjectContext projectContext)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.workbook = workbook;
            context = projectContext;
            cells = new ExcelCellGateway(application);
            ocr = new TesseractOcrEngine();

            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);
            Dock = DockStyle.Fill;
            Size = new Size(760, 700);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9F);

            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96, 96);
            view = new WorkspaceView();
            Controls.Add(view);
            WireProjectActions();
            documents = view.Documents;
            searchBox = view.Query;
            searchResults = view.Results;
            ocrState = view.IndexState;
            modeState = view.ModeState;
            status = view.Status;
            cancelButton = view.Cancel;
            canvas = view.Canvas;
            WireAnnotationActions();
            documents.SelectedIndexChanged += Documents_SelectedIndexChanged;
            searchResults.SelectedIndexChanged += SearchResults_SelectedIndexChanged;
            view.Import.Click += (sender, args) => ImportDocuments();
            view.Search.Click += (sender, args) => SearchFromPane();
            searchBox.KeyDown += (sender, args) => {
                if (args.KeyCode == Keys.Enter) { args.SuppressKeyPress = true; SearchFromPane(); }
                if (args.KeyCode == Keys.Escape) { args.SuppressKeyPress = true; view.HideResults(); }
            };
            view.Proofs.SelectedIndexChanged += (sender, args) => {
                if (bindingProofs || context.IsBusy) return;
                var item = view.Proofs.SelectedItem as ProofItem;
                if (item == null) return;
                try { EnsureActiveWorkbook(); NavigateToSnip(item.Id); }
                catch (Exception exception) { SetStatus(exception.Message); }
            };
            canvas.SelectionCompleted += Canvas_SelectionCompleted;
            cancelButton.Click += (sender, args) => operation?.Cancel();
            var reindexButton = view.Reindex;
            reindexButton.Click += async (sender, args) =>
            {
                if (context.IsBusy) return;
                try
                {
                    EnsureProject(); BeginOperation();
                    foreach (var document in context.State.Documents) document.IndexComplete = false;
                    var errors = await IndexMissingAsync();
                    BindDocuments();
                    SetStatus(errors.Count == 0 ? "Index mis à jour." : errors.Count + " pièce(s) à vérifier.");
                    if (errors.Count > 0) MessageBox.Show(this, string.Join("\n", errors), "Indexation");
                }
                catch (OperationCanceledException) { SetStatus("Réindexation annulée."); }
                catch (Exception exception) { ShowError(exception); }
                finally { EndOperation(); }
            };

            var removeButton = view.Remove;
            removeButton.Click += (sender, args) =>
            {
                if (context.IsBusy) return;
                try
                {
                    EnsureProject();
                    var document = SelectedDocument;
                    if (document == null) return;
                    if (context.State.Snips.Any(snip => snip.DocumentId == document.Id))
                        throw new InvalidOperationException("Cette pièce est liée à des preuves et doit être conservée.");
                    if (MessageBox.Show(this, "Retirer " + document.OriginalName + " de la liste ?", "Retirer une pièce", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                    var index = context.State.Documents.IndexOf(document);
                    var entry = new AuditEventRecord { Actor = Environment.UserName, Action = "DocumentRemoved", EntityType = "Document", EntityId = document.Id, Details = document.OriginalName };
                    context.State.Documents.Remove(document); context.State.AuditTrail.Add(entry);
                    try { context.Store.Save(context.State); }
                    catch { context.State.Documents.Insert(index, document); context.State.AuditTrail.Remove(entry); throw; }
                    searchResults.DataSource = null;
                    BindDocuments();
                    SetStatus("Pièce retirée de la liste. Sa copie locale est conservée pour récupération.");
                }
                catch (Exception exception) { ShowError(exception); }
            };

        }

        public void RefreshProject()
        {
            if (context.IsBusy || IsDisposed) return;
            try { EnsureProject(); RefreshCategories(); BindDocuments(); if(context.Store.RecoveryNotice!=null)SetStatus(context.Store.RecoveryNotice); }
            catch (Exception exception) { SetStatus(exception.Message); }
        }

        public void SetSnipMode(SnipType? type)
        {
            if(type.HasValue)view.SetReadingMode(false);
            activeSnipType = type;
            view.SetMode(type);
            Ribbon.DoctrackerRibbon.Instance?.Refresh();
            SetStatus(type.HasValue
                ? "Mode " + type.Value + " activé. Vous pouvez sniper plusieurs zones sans recliquer."
                : "Mode de snip désactivé.");
        }

        public bool IsSnipMode(SnipType type)
        {
            return activeSnipType.HasValue && activeSnipType.Value == type;
        }

        public void ImportDocuments()
        {
            ImportDocumentsAsync();
        }

        private async void ImportDocumentsAsync()
        {
            if (context.IsBusy) { SetStatus("Une opération est déjà en cours."); return; }
            try
            {
                EnsureProject();
                using (var dialog = new OpenFileDialog
                {
                    Title = "Ajouter des pièces au dossier Doctracker",
                    Filter = "Documents|*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp;*.doc;*.docx", Multiselect = true, CheckFileExists = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    BeginOperation();
                    lastImportedId=null;var errors = new List<string>();
                    foreach (var path in dialog.FileNames)
                    {
                        operation.Token.ThrowIfCancellationRequested();
                        try { SetStatus("Import : " + Path.GetFileName(path)); var imported=await Task.Run(() => WordDocumentImporter.Import(context,path,null,true,operation.Token));lastImportedId=imported.Id; }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception exception) { errors.Add(Path.GetFileName(path) + " : " + exception.Message); }
                    }
                    RefreshCategories(); BindDocuments();SelectLastImported();
                    errors.AddRange(await IndexMissingAsync());
                    BindDocuments();SelectLastImported();
                    SetStatus(errors.Count == 0 ? "Pièces importées et prêtes pour la recherche." : "Import terminé avec " + errors.Count + " erreur(s).");
                    if (errors.Count > 0) MessageBox.Show(this, string.Join("\n", errors), "Pièces à vérifier");
                }
            }
            catch (OperationCanceledException) { RefreshCategories();BindDocuments();SelectLastImported();SetStatus("Import interrompu. Les pièces déjà importées sont conservées."); }
            catch (Exception exception) { ShowError(exception); }
            finally { EndOperation(); }
        }

        private async void Canvas_SelectionCompleted(object sender, EventArgs e)
        {
            if(context.IsBusy)return;
            if(canvas.CommentMode){CreateDocumentComment();return;}
            if (!activeSnipType.HasValue) return;
            await CaptureSnipAsync(activeSnipType.Value);
        }

        private async Task CaptureSnipAsync(SnipType type)
        {
            if (context.IsBusy) return;
            try
            {
                EnsureProject();
                var document = SelectedDocument;
                if (document == null || !canvas.HasSelection) throw new InvalidOperationException("Dessinez une zone sur une pièce.");
                var target = cells.GetSingleTarget();
                var pageNumber = canvas.CurrentPageNumber;
                var rectangle = canvas.GetNormalizedSelection();
                BeginOperation();
                PageTextRecord recognized;
                if (type == SnipType.Validation || type == SnipType.Exception)
                    recognized = new PageTextRecord { Text = string.IsNullOrWhiteSpace(ExcelCellGateway.QueryText(target)) ? SnipTheme.LabelFor(type) : ExcelCellGateway.QueryText(target) };
                else
                {
                    SetStatus("Extraction de la zone…");
                    // Native/indexed words preserve exact text; scan-only pages use the local OCR.
                    recognized = ExtractIndexedSelection(document, pageNumber, rectangle);
                    if (recognized == null)
                        using (var crop = canvas.CropSelection()) recognized = await Task.Run(() => ocr.Recognize(crop, type == SnipType.Table));
                }
                operation.Token.ThrowIfCancellationRequested();
                EnsureActiveWorkbook();
                var writes = new List<PendingWrite>();
                if (type == SnipType.Table)
                {
                    var extracted = LayoutExtractor.Extract(recognized.Words);
                    if (extracted.Count == 0) throw new InvalidOperationException("Aucun tableau reconnu. Agrandissez la zone.");
                    var rowCount = extracted.Max(cell => cell.Row) + 1;
                    var columnCount = extracted.Max(cell => cell.Column) + 1;
                    if (rowCount * columnCount > 10000) throw new InvalidOperationException("Limitez le tableau à 10 000 cellules.");
                    if (target.Row + rowCount - 1 > target.Worksheet.Rows.Count || target.Column + columnCount - 1 > target.Worksheet.Columns.Count)
                        throw new InvalidOperationException("Le tableau dépasse les limites de la feuille.");
                    using (var preview = new TablePreviewDialog(extracted, rowCount, columnCount))
                    {
                        if (preview.ShowDialog(this) != DialogResult.OK) return;
                        for (var row = 0; row < rowCount; row++)
                        for (var column = 0; column < columnCount; column++)
                        {
                            var text = preview.ValueAt(row, column);
                            if (string.IsNullOrWhiteSpace(text)) continue;
                            var source = extracted.FirstOrDefault(cell => cell.Row == row && cell.Column == column);
                            var zone = source == null ? rectangle : new RectangleF(
                                rectangle.X + (float)source.X * rectangle.Width, rectangle.Y + (float)source.Y * rectangle.Height,
                                (float)source.Width * rectangle.Width, (float)source.Height * rectangle.Height);
                            var destination = (ExcelInterop.Range)target.Offset[row, column];
                            var cellType = InferType(text);
                            var write = PrepareWrite(destination, document, pageNumber, zone, cellType, text);
                            write.Snip.SourceType = SnipType.Table;
                            if (source != null && source.Text != text) write.Snip.Comment = "Texte corrigé dans l'aperçu. OCR : " + source.Text;
                            writes.Add(write);
                        }
                    }
                }
                else writes.Add(PrepareWrite(target, document, pageNumber, rectangle, type, recognized.Text));
                var preserveValue = type == SnipType.Validation || type == SnipType.Exception;
                var appendSum = false;
                if (type == SnipType.Sum)
                {
                    var previous = cells.GetSnipIds(target).Select(id => context.State.Snips.FirstOrDefault(snip => snip.Id == id)).ToList();
                    if (previous.Count > 0 && previous.All(snip => snip != null && snip.Type == SnipType.Sum))
                    {
                        var total = previous.Sum(snip => decimal.Parse(snip.ExtractedValue, System.Globalization.CultureInfo.InvariantCulture));
                        decimal cellValue;
                        if (!decimal.TryParse(Convert.ToString(target.Value2, System.Globalization.CultureInfo.InvariantCulture),
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out cellValue) || cellValue != total)
                            throw new InvalidOperationException("La somme a été modifiée dans Excel. Choisissez une cellule vide pour une nouvelle somme.");
                        writes[0].NumericOverride = total + decimal.Parse(writes[0].Snip.ExtractedValue, System.Globalization.CultureInfo.InvariantCulture);
                        appendSum = true;
                    }
                }
                if (!ConfirmOverwrite(writes, preserveValue || appendSum)) return;
                CommitWrites(writes, preserveValue || appendSum);
                canvas.ClearSelection();
                UpdateDocumentProofs();
                SetStatus(writes.Count + " preuve(s) créée(s). Sélectionnez la prochaine cellule ou dessinez la zone suivante.");
                if (writes.Count > 0 && !preserveValue && type != SnipType.Sum && writes.Max(w => w.Target.Row) < target.Worksheet.Rows.Count &&
                    (application.Selection as ExcelInterop.Range)?.Cells.CountLarge == 1 &&
                    cells.GetSingleTarget().Address[true, true, ExcelInterop.XlReferenceStyle.xlA1, true] ==
                    target.Address[true, true, ExcelInterop.XlReferenceStyle.xlA1, true])
                    ((ExcelInterop.Range)target.Offset[type == SnipType.Table ? writes.Max(w => w.Target.Row) - target.Row + 1 : 1, 0]).Select();
            }
            catch (OperationCanceledException) { SetStatus("Extraction annulée."); }
            catch (Exception exception) { ShowError(exception); }
            finally { EndOperation(); }
        }

        public async void SearchSelection()
        {
            try
            {
                var target = cells.GetSingleTarget();
                var query = ExcelCellGateway.QueryText(target,true);
                if (string.IsNullOrWhiteSpace(query))
                    throw new InvalidOperationException("La cellule active est vide.");
                searchBox.Text = query;
                await SearchDocumentsAsync(query);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public async void SearchFromPane()
        {
            if(string.IsNullOrWhiteSpace(searchBox.Text)){SearchSelection();return;}
            await SearchDocumentsAsync(searchBox.Text);
        }

        private async Task SearchDocumentsAsync(string query)
        {
            if (context.IsBusy) { SetStatus("Une opération est déjà en cours."); return; }
            try
            {
                EnsureProject();
                if (string.IsNullOrWhiteSpace(query)) throw new InvalidOperationException("Saisissez un texte à rechercher.");
                BeginOperation();
                var errors = await IndexMissingAsync();
                var scope=SearchScope();scope.Documents=scope.Documents.Where(d=>d.IndexComplete).ToList();
                var results = await Task.Run(() => OccurrenceSearch.Find(scope, query, 201, operation.Token)
                    .Select(candidate => new SearchResultItem { Candidate = candidate,
                        Document = context.State.Documents.First(item => item.Id == candidate.DocumentId) }).ToList());
                operation.Token.ThrowIfCancellationRequested();
                var truncated=results.Count>200;if(truncated)results=results.Take(200).ToList();
                searchResults.DataSource = results;
                view.ShowResults(results.Count);
                SetStatus((results.Count==0 ? "Aucun résultat. Vérifiez le dossier ou utilisez Documents > Réindexer par OCR dans le ruban." : results.Count + " occurrence(s)."+(truncated?" Affichage limité à 200 : précisez la recherche.":"")) + (errors.Count > 0 ? " Attention : " + errors.Count + " pièce(s) non indexée(s)." : ""));
            }
            catch (OperationCanceledException) { SetStatus("Recherche annulée."); }
            catch (Exception exception) { ShowError(exception); }
            finally { EndOperation(); }
        }

        public void SetMatchingInputSelection()
        {
            try
            {
                EnsureProject();
                if (context.IsBusy) return;
                var range = cells.GetSelection();
                ValidateMatchingColumn(range, "recherche");
                matchingInputSheet = ((ExcelInterop.Worksheet)range.Worksheet).Name;
                matchingInputAddress = range.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1];
                SetStatus("Colonne de recherche définie : " + matchingInputSheet + "!" + matchingInputAddress);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void SetMatchingOutputSelection()
        {
            try
            {
                EnsureProject();
                if (context.IsBusy) return;
                var range = cells.GetSelection();
                ValidateMatchingColumn(range, "résultat");
                matchingOutputSheet = ((ExcelInterop.Worksheet)range.Worksheet).Name;
                matchingOutputAddress = range.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1];
                SetStatus("Colonne de résultat définie : " + matchingOutputSheet + "!" + matchingOutputAddress);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public async void MatchSelection()
        {
            if (context.IsBusy) { SetStatus("Une opération est déjà en cours."); return; }
            try
            {
                EnsureProject();
                var input = ResolveMatchingRange(matchingInputSheet, matchingInputAddress);
                var output = ResolveMatchingRange(matchingOutputSheet, matchingOutputAddress);
                if (input == null || output == null) throw new InvalidOperationException("Définissez la plage de recherche puis la destination.");
                var rowCount = input.Rows.Count;
                var columnCount = input.Columns.Count;
                ValidateMatchingColumn(input, "recherche");
                if (output.Cells.CountLarge != 1 && (output.Rows.Count < rowCount || output.Columns.Count < columnCount))
                    throw new InvalidOperationException("La destination doit couvrir toutes les lignes et colonnes (ou être une cellule de départ).");
                var first = (ExcelInterop.Range)output.Cells[1, 1];
                if (first.Row + rowCount - 1 > first.Worksheet.Rows.Count || first.Column + columnCount - 1 > first.Worksheet.Columns.Count)
                    throw new InvalidOperationException("La destination dépasse la feuille.");
                var destination = first.Resize[rowCount, columnCount];
                ExcelCellGateway.ValidateWritable(destination);
                if (input.Worksheet.Name == destination.Worksheet.Name && application.Intersect(input, destination) != null)
                    throw new InvalidOperationException("Les plages de recherche et de résultat ne doivent pas se chevaucher.");
                // Snapshot all inputs before yielding: worksheet edits cannot change this run's criteria.
                var queries = new List<string[]>();
                for (var row = 1; row <= rowCount; row++)
                    queries.Add(Enumerable.Range(1, columnCount).Select(column => ExcelCellGateway.QueryText((ExcelInterop.Range)input.Cells[row, column])).ToArray());
                BeginOperation();
                var errors = await IndexMissingAsync();
                if (errors.Count > 0) throw new InvalidOperationException("Matching interrompu : certaines pièces ne sont pas indexées.\n" + string.Join("\n", errors));
                var scope=SearchScope();scope.Documents=scope.Documents.Where(d=>d.IndexComplete).ToList();
                var results = await Task.Run(() => context.Matcher.FindBatch(scope,queries.Select(q=>(IReadOnlyList<string>)q).ToList(),true,operation.Token));
                operation.Token.ThrowIfCancellationRequested();
                EnsureActiveWorkbook();
                var writes = new List<PendingWrite>();
                var ambiguous = 0;
                var missing = 0;
                for (var row = 0; row < rowCount; row++)
                {
                    if (queries[row].All(string.IsNullOrWhiteSpace)) continue;
                    var candidates = results[row];
                    if (candidates.Count == 0) { missing++; continue; }
                    if (candidates.Count > 1) { ambiguous++; continue; }
                    var candidate = candidates[0];
                    var document = context.State.Documents.First(item => item.Id == candidate.DocumentId);
                    for (var column = 0; column < columnCount; column++)
                    {
                        var field = candidate.Fields[column];
                        if (field == null) continue;
                        var target = (ExcelInterop.Range)first.Offset[row, column];
                        var zone = new RectangleF((float)field.X, (float)field.Y, (float)field.Width, (float)field.Height);
                        var write = PrepareWrite(target, document, candidate.PageNumber, zone, InferType(field.Evidence), field.Evidence);
                        write.Snip.Comment = (candidate.IsPartial ? "Rapprochement par référence partielle, à vérifier." : "Rapprochement exact, à revoir.") + (field.HasLocation ? "" : " Localisation : page entière.");
                        writes.Add(write);
                    }
                }
                // A re-run must not leave stale previous matches in unresolved rows.
                var unresolved = new List<ExcelInterop.Range>();
                for (var row = 0; row < rowCount; row++)
                for (var column = 0; column < columnCount; column++)
                    if (results[row].Count != 1 || string.IsNullOrWhiteSpace(queries[row][column]))
                        unresolved.Add((ExcelInterop.Range)first.Offset[row, column]);
                var occupied = writes.Any(write => ExcelCellGateway.HasContent(write.Target)) || unresolved.Any(ExcelCellGateway.HasContent);
                if (occupied && MessageBox.Show(this, "La destination contient des données. Remplacer les résultats et vider les lignes sans correspondance ?", "Matching", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                if(results.Any(row=>row.Count==1 && row[0].IsPartial) && MessageBox.Show(this,"Des références partielles ont été trouvées. Insérer les valeurs réellement lues et leurs preuves pour revue ?","Références partielles",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                CommitWrites(writes, false, unresolved);
                SetStatus(writes.Count + " preuve(s) créée(s) ; " + missing + " ligne(s) sans résultat ; " + ambiguous + " ambiguë(s). Revue requise.");
                if (missing + ambiguous > 0) MessageBox.Show(this,
                    "Lignes sans correspondance : " + string.Join(", ", Enumerable.Range(0, rowCount).Where(i => results[i].Count == 0 && queries[i].Any(q => !string.IsNullOrWhiteSpace(q))).Select(i => input.Row + i)) +
                    "\nLignes ambiguës : " + string.Join(", ", Enumerable.Range(0, rowCount).Where(i => results[i].Count > 1).Select(i => input.Row + i)), "Résultats à compléter");
            }
            catch (OperationCanceledException) { SetStatus("Matching annulé avant insertion."); }
            catch (Exception exception) { ShowError(exception); }
            finally { EndOperation(); }
        }

        // Passive selection navigation never opens a dialog or changes the Excel selection.
        public bool TryNavigateFromCell(ExcelInterop.Range target)
        {
            try
            {
                if (context.IsBusy) return false;
                var ids = cells.GetSnipIds(target);
                if (ids.Count == 0) { ClearCellProof(); return false; }
                EnsureProject();
                var items = ids.Select(id => context.State.Snips.FirstOrDefault(snip => snip.Id == id))
                    .Where(snip => snip != null).Select(snip => new ProofItem { Id = snip.Id,
                        Caption = SnipTheme.LabelFor(snip.SourceType ?? snip.Type) + " · " +
                        context.State.Documents.FirstOrDefault(doc => doc.Id == snip.DocumentId)?.OriginalName +
                        " · p. " + snip.PageNumber + " · " + snip.WorksheetName + "!" + snip.CellAddress }).ToList();
                if (items.Count == 0) { ClearCellProof(); return false; }
                bindingProofs = true;
                try { view.Proofs.DataSource = items; view.Proofs.SelectedIndex = items.Count - 1; view.ShowProofs(true); }
                finally { bindingProofs = false; }
                return NavigateToSnip(items[items.Count - 1].Id);
            }
            catch (Exception exception) { SetStatus("Preuve indisponible : " + exception.Message); return false; }
        }

        public void ClearCellProof()
        {
            if (context.IsBusy) return;
            focusedSnipId = null;
            view.ShowProofs(false);
            canvas.ClearSelection();
        }

        public void NavigateFromSelection()
        {
            try
            {
                if (context.IsBusy) return;
                EnsureProject();
                var target = cells.GetSingleTarget();
                var snipId = ChooseSnipId(target);
                if (string.IsNullOrWhiteSpace(snipId))
                    throw new InvalidOperationException("La cellule sélectionnée n'a aucune preuve Doctracker.");
                NavigateToSnip(snipId);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void ReviewSelection()
        {
            try
            {
                if (context.IsBusy) return;
                EnsureProject();
                var target = cells.GetSingleTarget();
                var snipId = ChooseSnipId(target);
                var snip = context.State.Snips.FirstOrDefault(item => item.Id == snipId);
                if (snip == null) throw new InvalidOperationException("La cellule sélectionnée n'a aucune preuve Doctracker.");

                using (var dialog = new ReviewDialog(snip))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    EnsureActiveWorkbook();
                    ExcelCellGateway.ValidateWritable(target);
                    var snapshot = cells.Snapshot(target);
                    var source = context.State.Documents.First(item => item.Id == snip.DocumentId);
                    try
                    {
                        context.Snips.SetReview(context.State, snip.Id, dialog.SelectedStatus,
                            dialog.ReviewComment, Environment.UserName, updated => cells.AttachProof(target, updated, source));
                    }
                    catch { snapshot.Restore(); throw; }
                }

                SetStatus("Statut de revue mis à jour : " + snip.Status + ".");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private bool NavigateToSnip(string snipId)
        {
            var snip = context.State.Snips.FirstOrDefault(item => item.Id == snipId);
            if (snip == null) return false;
            var document = context.State.Documents.FirstOrDefault(item => item.Id == snip.DocumentId);
            if (document == null) return false;

            BindDocuments();
            SelectDocument(document.Id);
            UpdateDocumentProofs();
            canvas.NavigateTo(context.Store.ResolveDocumentPath(document), snip);
            focusedSnipId = snip.Id;
            SetStatus(document.OriginalName + " — page " + snip.PageNumber + " — " + snip.Status);
            return true;
        }

        private void UpdateDocumentProofs()
        {
            var document = SelectedDocument;
            canvas.SetDocument(document);
            canvas.SetProofs(document == null ? Enumerable.Empty<SnipRecord>() : context.State.Snips.Where(snip => snip.DocumentId == document.Id));
        }

        private void Documents_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var document = SelectedDocument;
                if (document != null && context.State.Documents.Any(item => item.Id == document.Id))
                {
                    if (!string.Equals(canvas.CurrentPath, context.Store.ResolveDocumentPath(document), StringComparison.OrdinalIgnoreCase)) canvas.LoadDocument(context.Store.ResolveDocumentPath(document));
                    UpdateDocumentProofs();
                    SetStatus(document.OriginalName + " — sélectionnez une zone ou recherchez dans les pièces.");
                }
                else canvas.ClearDocument();
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private void SearchResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var result = searchResults.SelectedItem as SearchResultItem;
                if (result == null) return;
                SelectDocument(result.Document.Id);
                canvas.NavigateTo(context.Store.ResolveDocumentPath(result.Document), new SnipRecord {
                    PageNumber = result.Candidate.PageNumber, X = result.Candidate.X, Y = result.Candidate.Y,
                    Width = result.Candidate.Width, Height = result.Candidate.Height });
                SetStatus(result.Document.DisplayName + " — page " + result.Candidate.PageNumber +
                          " — " + (result.Candidate.HasLocation ? "occurrence localisée" : "texte trouvé ; position non disponible"));
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private DocumentRecord SelectedDocument => documents.SelectedItem as DocumentRecord;

        private void EnsureProject()
        {
            EnsureActiveWorkbook();
            context.Ensure(workbook);
            if (boundProjectPath != context.WorkbookPath)
            {
                boundProjectPath = context.WorkbookPath;
                matchingInputSheet = matchingInputAddress = matchingOutputSheet = matchingOutputAddress = null;
                searchResults.DataSource = null;
                view.HideResults(); view.ShowProofs(false); focusedSnipId = null;
                canvas.ClearDocument();
                BindDocuments();
            }
        }

        private void BindDocuments()
        {
            var visible=VisibleDocuments().Reverse().OrderByDescending(d=>d.LastImportedAtUtc==default(DateTime)?d.AddedAtUtc:d.LastImportedAtUtc).ToList();
            var selectedId = SelectedDocument?.Id;
            if(selectedId!=null && !visible.Any(d=>d.Id==selectedId))selectedId=null;
            var listChanged = documents.Items.Count != visible.Count ||
                documents.Items.Cast<DocumentRecord>().Where((doc, index) => !ReferenceEquals(doc, visible[index])).Any();
            if (listChanged)
            {
                documents.SelectedIndexChanged -= Documents_SelectedIndexChanged;
                documents.DataSource = visible;
                documents.DisplayMember = "DisplayName";
                if (selectedId != null) SelectDocument(selectedId);
                if (documents.SelectedIndex < 0 && documents.Items.Count > 0) documents.SelectedIndex = 0;
                documents.SelectedIndexChanged += Documents_SelectedIndexChanged;
                Documents_SelectedIndexChanged(this, EventArgs.Empty);
            }
            var total = context.State.Documents.Count;
            var indexed = context.State.Documents.Count(item => item.IndexComplete);
            ocrState.Text = total == 0
                ? "Aucune pièce"
                : total + " pièces · " + indexed + " indexées";
        }

        private void SelectLastImported() { if(lastImportedId!=null)SelectDocument(lastImportedId); }

        private void SelectDocument(string id)
        {
            if(!documents.Items.Cast<DocumentRecord>().Any(d=>d.Id==id) && context.State.Documents.Any(d=>d.Id==id))
            {
                bindingCategories=true;view.Categories.SelectedIndex=0;bindingCategories=false;BindDocuments();
            }
            for (var index = 0; index < documents.Items.Count; index++)
            {
                var document = documents.Items[index] as DocumentRecord;
                if (document != null && document.Id == id)
                {
                    documents.SelectedIndex = index;
                    return;
                }
            }
        }

        private static void ValidateMatchingColumn(ExcelInterop.Range range, string purpose)
        {
            if (range == null || range.Areas.Count != 1 || range.Columns.Count > 10)
                throw new InvalidOperationException("Sélectionnez une plage continue de 1 à 10 colonnes pour le " + purpose + ".");
            if (range.Rows.Count > 5000)
                throw new InvalidOperationException("Limitez la sélection de " + purpose + " à 5 000 lignes.");
        }

        private ExcelInterop.Range ResolveMatchingRange(string worksheetName, string address)
        {
            if (string.IsNullOrWhiteSpace(worksheetName) || string.IsNullOrWhiteSpace(address)) return null;

            var worksheet = workbook.Worksheets[worksheetName] as ExcelInterop.Worksheet;
            return worksheet == null ? null : worksheet.Range[address];
        }

        private void SetStatus(string message)
        {
            if (!IsDisposed) status.Text = message;
        }

        private void SetStatusThreadSafe(string message)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action<string>(SetStatus), message); } catch (InvalidOperationException) { }
            }
            else
            {
                SetStatus(message);
            }
        }

        private void ShowError(Exception exception)
        {
            if (IsDisposed) return;
            var detail=exception.GetBaseException().Message;
            var message=exception.Message+(detail==exception.Message?"":"\n"+detail);
            SetStatus("Erreur : " + message);
            MessageBox.Show(this, message, "Doctracker",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void EnsureActiveWorkbook()
        {
            if (application.ActiveWorkbook == null || !string.Equals(application.ActiveWorkbook.FullName, workbook.FullName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Revenez au classeur de cette opération.");
            if (operation != null && !string.Equals(workbook.FullName, context.WorkbookPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Le classeur a changé de nom pendant l'opération. Relancez après actualisation.");
        }

        private void BeginOperation()
        {
            if (context.IsBusy) throw new InvalidOperationException("Une opération est déjà en cours dans ce classeur.");
            context.IsBusy = true;
            operation = new CancellationTokenSource();
            canvas.Enabled = documents.Enabled = searchResults.Enabled = false;
            view.SetBusy(true);
        }

        private void EndOperation()
        {
            if (operation == null) return;
            operation.Dispose(); operation = null; context.IsBusy = false;
            if (IsDisposed) return;
            canvas.Enabled = documents.Enabled = searchResults.Enabled = true;
            view.SetBusy(false);
            context.MarkWorkbookDirty();
        }

        private Task<List<string>> IndexMissingAsync()
        {
            var indexer = new DocumentIndexer(context.Store, ocr);
            var token = operation.Token;
            return Task.Run(() => indexer.IndexMissing(context.State,
                (name, page, count) => SetStatusThreadSafe("Indexation : " + name + " — " + page + "/" + count), token));
        }

        private PendingWrite PrepareWrite(ExcelInterop.Range target, DocumentRecord document, int page, RectangleF zone, SnipType type, string text)
        {
            var snip = context.Snips.Prepare(context.State, document.Id, page,
                new NormalizedRectangle(zone.X, zone.Y, zone.Width, zone.Height), type, text,
                target.Worksheet.Name, target.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1], Environment.UserName);
            return new PendingWrite { Target = target, Snip = snip, Document = document };
        }

        private bool ConfirmOverwrite(List<PendingWrite> writes, bool preserveValue)
        {
            foreach (var write in writes) ExcelCellGateway.ValidateWritable(write.Target);
            return preserveValue || !writes.Any(write => ExcelCellGateway.HasContent(write.Target)) ||
                MessageBox.Show(this, "La destination contient déjà des données. Les remplacer par l'extraction ?", "Confirmer l'insertion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private void CommitWrites(List<PendingWrite> writes, bool append, List<ExcelInterop.Range> clear = null)
        {
            EnsureActiveWorkbook();
            operation?.Token.ThrowIfCancellationRequested();
            var targets = writes.Select(write => write.Target).Concat(clear ?? new List<ExcelInterop.Range>()).ToList();
            foreach (var target in targets) ExcelCellGateway.ValidateWritable(target);
            var snapshots = targets.Select(cells.Snapshot).ToList();
            var events = application.EnableEvents;
            var updating = application.ScreenUpdating;
            try
            {
                application.EnableEvents = false; application.ScreenUpdating = false;
                foreach (var target in clear ?? new List<ExcelInterop.Range>())
                {
                    target.ClearContents();
                    cells.RemoveProof(target);
                }
                foreach (var write in writes)
                {
                    cells.WriteSnip(write.Target, write.Snip, write.Document, append);
                    if (write.NumericOverride.HasValue) write.Target.Value2 = (double)write.NumericOverride.Value;
                }
                context.Snips.Commit(context.State, writes.Select(write => write.Snip).ToList());
            }
            catch (Exception failure)
            {
                var failures = new List<Exception> { failure };
                foreach (var snapshot in snapshots)
                    try { snapshot.Restore(); } catch (Exception rollback) { failures.Add(rollback); }
                if (failures.Count > 1) throw new AggregateException("Insertion échouée et restauration incomplète. Vérifiez la destination avant de continuer.", failures);
                throw;
            }
            finally { application.EnableEvents = events; application.ScreenUpdating = updating; }
        }

        private static SnipType InferType(string text)
        {
            var value = (text ?? "").Trim();
            // Leading zero identifiers remain text.
            decimal number;
            if (!System.Text.RegularExpressions.Regex.IsMatch(value, @"^0\d") && new TextValueParser().TryParseAmount(value, out number)) return SnipType.Number;
            if (System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$"))
                try { TextValueParser.ParseDate(value); return SnipType.Date; } catch (FormatException) { }
            return SnipType.Text;
        }

        private static PageTextRecord ExtractIndexedSelection(DocumentRecord document, int pageNumber, RectangleF zone)
        {
            var page = document.IndexedPages.FirstOrDefault(item => item.PageNumber == pageNumber);
            if (page == null || page.Words == null || page.Words.Count == 0) return null;
            var selected = page.Words.Where(word =>
                word.X + word.Width / 2 >= zone.Left && word.X + word.Width / 2 <= zone.Right &&
                word.Y + word.Height / 2 >= zone.Top && word.Y + word.Height / 2 <= zone.Bottom).ToList();
            if (selected.Count == 0) return null;
            return new PageTextRecord
            {
                Text = string.Join("\n", selected.GroupBy(word => word.Line).Select(line => string.Join(" ", line.Select(word => word.Text)))),
                Words = selected.Select(word => new WordRecord { Text = word.Text, Line = word.Line,
                    X = Math.Max(0, (word.X - zone.X) / zone.Width), Y = Math.Max(0, (word.Y - zone.Y) / zone.Height),
                    Width = Math.Min(word.X + word.Width, zone.Right) / zone.Width - Math.Max(word.X, zone.Left) / zone.Width,
                    Height = Math.Min(word.Y + word.Height, zone.Bottom) / zone.Height - Math.Max(word.Y, zone.Top) / zone.Height }).ToList()
            };
        }

        private string ChooseSnipId(ExcelInterop.Range target)
        {
            var ids = cells.GetSnipIds(target);
            if (focusedSnipId != null && ids.Contains(focusedSnipId)) return focusedSnipId;
            if (ids.Count < 2) return ids.FirstOrDefault();
            using (var dialog = new Form { Text = "Choisir la preuve", Width = 540, Height = 260, StartPosition = FormStartPosition.CenterParent })
            using (var list = new ListBox { Dock = DockStyle.Fill, DisplayMember = "Caption" })
            using (var button = new Button { Text = "Ouvrir", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK })
            {
                var items = ids.Select(id => context.State.Snips.FirstOrDefault(snip => snip.Id == id)).Where(snip => snip != null)
                    .Select(snip => new { snip.Id, Caption = snip.Type + " — " + context.State.Documents.FirstOrDefault(doc => doc.Id == snip.DocumentId)?.OriginalName + " — page " + snip.PageNumber }).ToList();
                list.DataSource = items;
                dialog.Controls.Add(list); dialog.Controls.Add(button); dialog.AcceptButton = button;
                if (dialog.ShowDialog(this) != DialogResult.OK || list.SelectedIndex < 0) return null;
                return items[list.SelectedIndex].Id;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                operation?.Cancel();
                // Disposal waits for native OCR use to finish; no COM work runs on the worker.
                ocr.Dispose();
            }
            base.Dispose(disposing);
        }

        private sealed class ProofItem
        {
            public string Id { get; set; }
            public string Caption { get; set; }
        }

        private sealed class PendingWrite
        {
            public decimal? NumericOverride { get; set; }
            public ExcelInterop.Range Target { get; set; }
            public SnipRecord Snip { get; set; }
            public DocumentRecord Document { get; set; }
        }

        private sealed class SearchResultItem
        {
            public MatchCandidate Candidate { get; set; }
            public DocumentRecord Document { get; set; }
            public string Caption => Document.DisplayName + " — page " + Candidate.PageNumber +
                                     " — " + Candidate.Evidence;
        }
    }
}
