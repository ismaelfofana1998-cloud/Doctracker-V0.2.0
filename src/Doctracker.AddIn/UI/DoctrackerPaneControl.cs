using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Doctracker.AddIn.Excel;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed class DoctrackerPaneControl : UserControl
    {
        private readonly ExcelInterop.Application application;
        private readonly WorkbookProjectContext context;
        private readonly ExcelCellGateway cells;
        private readonly IOcrEngine ocr;
        private readonly DocumentCanvas canvas;
        private readonly ComboBox documents;
        private TextBox searchBox;
        private readonly ListBox searchResults;
        private readonly Label ocrState;
        private Label modeState;
        private readonly Label status;

        private SnipType? activeSnipType;
        private bool snipInProgress;
        private string matchingInputSheet;
        private string matchingInputAddress;
        private string matchingOutputSheet;
        private string matchingOutputAddress;

        public DoctrackerPaneControl(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            context = new WorkbookProjectContext();
            cells = new ExcelCellGateway(application);
            ocr = new TesseractOcrEngine();

            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);
            Dock = DockStyle.Fill;
            Size = new Size(760, 700);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9F);

            var header = BuildHeader();
            var searchBar = BuildSearchBar(out searchBox);
            var workflowBar = BuildWorkflowBar(out modeState);

            documents = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = "OriginalName",
                FlatStyle = FlatStyle.Flat,
                IntegralHeight = false,
                BackColor = Color.White
            };
            documents.SelectedIndexChanged += Documents_SelectedIndexChanged;

            var importButton = new Button
            {
                Text = "Ajouter des pièces — OCR automatique",
                Dock = DockStyle.Top,
                Height = 34,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(18, 28, 45),
                ForeColor = Color.White,
                TabStop = false
            };
            importButton.FlatAppearance.BorderSize = 0;
            importButton.Click += (sender, args) => ImportDocuments();

            ocrState = new Label
            {
                Text = "Aucune pièce",
                Dock = DockStyle.Top,
                Height = 24,
                Padding = new Padding(6, 4, 0, 0),
                ForeColor = Color.FromArgb(80, 88, 98),
                BackColor = Color.FromArgb(248, 249, 251)
            };

            searchResults = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                DisplayMember = "Caption"
            };
            searchResults.SelectedIndexChanged += SearchResults_SelectedIndexChanged;

            var documentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
                BackColor = Color.White
            };
            documentPanel.Controls.Add(searchResults);
            documentPanel.Controls.Add(ocrState);
            documentPanel.Controls.Add(importButton);
            documentPanel.Controls.Add(documents);

            canvas = new DocumentCanvas();
            canvas.SelectionCompleted += Canvas_SelectionCompleted;

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                FixedPanel = FixedPanel.Panel1,
                Panel1MinSize = 185,
                BackColor = Color.FromArgb(231, 235, 240)
            };
            split.SplitterDistance = 225;
            split.Panel1.Controls.Add(documentPanel);
            split.Panel2.Controls.Add(canvas);

            status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Text = "Prêt",
                Padding = new Padding(10, 6, 0, 0),
                ForeColor = Color.FromArgb(80, 88, 98),
                BackColor = Color.FromArgb(248, 249, 251),
                AutoEllipsis = true
            };

            Controls.Add(split);
            Controls.Add(status);
            Controls.Add(workflowBar);
            Controls.Add(searchBar);
            Controls.Add(header);
        }

        public void RefreshProject()
        {
            try
            {
                EnsureProject();
                BindDocuments();
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void SetSnipMode(SnipType? type)
        {
            activeSnipType = type;
            modeState.Text = type.HasValue
                ? "Mode actif : " + type.Value + " — dessinez les zones à la suite"
                : "Mode actif : aucun — choisissez un type dans le ruban";
            modeState.ForeColor = type.HasValue
                ? Color.FromArgb(24, 112, 70)
                : Color.FromArgb(80, 88, 98);
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
            try
            {
                EnsureProject();
                using (var dialog = new OpenFileDialog
                {
                    Title = "Ajouter des pièces au dossier Doctracker",
                    Filter = "Documents|*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp",
                    Multiselect = true,
                    CheckFileExists = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;

                    var imported = new List<DocumentRecord>();
                    foreach (var path in dialog.FileNames)
                    {
                        var document = context.Importer.Import(context.State, path, Environment.UserName);
                        if (!imported.Any(item => item.Id == document.Id)) imported.Add(document);
                    }

                    BindDocuments();
                    var indexer = new DocumentIndexer(context.Store, ocr);
                    foreach (var document in imported.Where(item => item.IndexedPages.Count == 0))
                    {
                        SetStatus("OCR automatique : " + document.OriginalName + "…");
                        await Task.Run(() => indexer.Index(context.State,
                            document,
                            (page, count) => SetStatusThreadSafe(
                                "OCR : " + document.OriginalName + " — page " + page + "/" + count)));
                    }
                }

                BindDocuments();
                SetStatus("Pièces importées et index OCR prêtes pour la recherche.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private async void Canvas_SelectionCompleted(object sender, EventArgs e)
        {
            if (!activeSnipType.HasValue || snipInProgress) return;
            await CaptureSnipAsync(activeSnipType.Value);
        }

        private async Task CaptureSnipAsync(SnipType type)
        {
            if (snipInProgress) return;
            snipInProgress = true;
            try
            {
                EnsureProject();
                var document = SelectedDocument;
                if (document == null) throw new InvalidOperationException("Sélectionnez d'abord une pièce.");
                if (!canvas.HasSelection) throw new InvalidOperationException("Dessinez une zone sur le document.");

                var target = cells.GetSingleTarget();
                var worksheet = (ExcelInterop.Worksheet)target.Worksheet;
                var address = target.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1];
                var rectangle = canvas.GetNormalizedSelection();
                string recognized;

                SetStatus("Reconnaissance OCR de la zone en cours…");
                using (var crop = canvas.CropSelection())
                {
                    recognized = await Task.Run(() => ocr.Recognize(crop));
                }

                var snip = context.Snips.Create(
                    context.State,
                    document.Id,
                    canvas.CurrentPageNumber,
                    new NormalizedRectangle(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height),
                    type,
                    recognized,
                    worksheet.Name,
                    address,
                    Environment.UserName);

                cells.WriteSnip(target, snip, document);
                canvas.ClearSelection();
                SetStatus(type + " créé dans " + worksheet.Name + "!" + address + ". Dessinez le suivant.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
            finally
            {
                snipInProgress = false;
            }
        }

        public async void SearchSelection()
        {
            try
            {
                var target = cells.GetSingleTarget();
                var query = Convert.ToString(target.Value2);
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

        public void SearchFromPane()
        {
            SearchDocumentsAsync(searchBox.Text);
        }

        private async Task SearchDocumentsAsync(string query)
        {
            try
            {
                EnsureProject();
                if (string.IsNullOrWhiteSpace(query))
                    throw new InvalidOperationException("Saisissez un texte ou placez-vous sur une cellule à rechercher.");

                var indexer = new DocumentIndexer(context.Store, ocr);
                if (context.State.Documents.Any(item => item.IndexedPages.Count == 0))
                {
                    SetStatus("Indexation OCR des pièces restantes…");
                    await Task.Run(() => indexer.IndexMissing(context.State,
                        (name, page, count) => SetStatusThreadSafe(
                            "OCR : " + name + " — page " + page + "/" + count)));
                }

                var results = context.Matcher.Find(context.State, query, 50)
                    .Select(candidate => new SearchResultItem
                    {
                        Candidate = candidate,
                        Document = context.State.Documents.First(item => item.Id == candidate.DocumentId)
                    })
                    .ToList();
                searchResults.DataSource = null;
                searchResults.DataSource = results;
                searchResults.DisplayMember = "Caption";
                SetStatus(results.Count == 0
                    ? "Aucun résultat dans les pièces indexées."
                    : results.Count + " résultat(s) trouvé(s) dans les pièces.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public void SetMatchingInputSelection()
        {
            try
            {
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
            try
            {
                EnsureProject();
                var input = ResolveMatchingRange(matchingInputSheet, matchingInputAddress);
                var output = ResolveMatchingRange(matchingOutputSheet, matchingOutputAddress);
                if (input == null || output == null)
                    throw new InvalidOperationException(
                        "Définissez d'abord la colonne de recherche puis la colonne de résultat dans le ruban.");
                if (input.Columns.Count != 1 || output.Columns.Count != 1)
                    throw new InvalidOperationException("Les deux sélections doivent contenir une seule colonne.");
                if (input.Rows.Count > 5000)
                    throw new InvalidOperationException("Limitez la plage de matching à 5 000 lignes.");
                if (output.Cells.CountLarge != 1 && output.Rows.Count < input.Rows.Count)
                    throw new InvalidOperationException("La colonne de résultat doit couvrir toutes les lignes de recherche.");

                SetStatus("Indexation OCR des pièces non indexées…");
                var indexer = new DocumentIndexer(context.Store, ocr);
                await Task.Run(() => indexer.IndexMissing(context.State,
                    (name, page, count) => SetStatusThreadSafe(
                        "Indexation : " + name + " — page " + page + "/" + count)));

                var matched = 0;
                var outputStart = (ExcelInterop.Range)output.Cells[1, 1];
                for (var row = 1; row <= input.Rows.Count; row++)
                {
                    var inputCell = (ExcelInterop.Range)input.Cells[row, 1];
                    var outputCell = output.Cells.CountLarge == 1
                        ? (ExcelInterop.Range)outputStart.Offset[row - 1, 0]
                        : (ExcelInterop.Range)output.Cells[row, 1];
                    var query = Convert.ToString(inputCell.Value2);
                    if (string.IsNullOrWhiteSpace(query))
                    {
                        outputCell.Value2 = string.Empty;
                        continue;
                    }

                    var candidate = context.Matcher.Find(context.State, query, 1).FirstOrDefault();
                    if (candidate == null || candidate.Score < 0.40)
                    {
                        outputCell.Value2 = "Aucun rapprochement";
                        continue;
                    }

                    var document = context.State.Documents.First(item => item.Id == candidate.DocumentId);
                    var worksheet = (ExcelInterop.Worksheet)outputCell.Worksheet;
                    var resultText = document.OriginalName + " — page " + candidate.PageNumber +
                                     " — score " + candidate.Score.ToString("P0");
                    outputCell.Value2 = resultText;
                    var snip = context.Snips.Create(
                        context.State,
                        document.Id,
                        candidate.PageNumber,
                        new NormalizedRectangle(0, 0, 1, 1),
                        SnipType.Text,
                        query,
                        worksheet.Name,
                        outputCell.Address[false, false, ExcelInterop.XlReferenceStyle.xlA1],
                        Environment.UserName);
                    snip.Comment = "Rapprochement automatique, score " +
                                   candidate.Score.ToString("P0") + ". " + candidate.Evidence;
                    context.Store.Save(context.State);
                    cells.AttachProof(outputCell, snip, document);
                    matched++;

                    if (row % 25 == 0) SetStatus("Matching : ligne " + row + "/" + input.Rows.Count + "…");
                }

                SetStatus(matched + " ligne(s) rapprochée(s) sur " + input.Rows.Count +
                          ". Résultats écrits dans la colonne de sortie ; validation humaine requise.");
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        public bool TryNavigateFromCell(ExcelInterop.Range target)
        {
            try
            {
                EnsureProject();
                var snipId = cells.GetSnipId(target);
                if (string.IsNullOrWhiteSpace(snipId)) return false;
                return NavigateToSnip(snipId);
            }
            catch (Exception exception)
            {
                ShowError(exception);
                return false;
            }
        }

        public void NavigateFromSelection()
        {
            try
            {
                EnsureProject();
                var target = cells.GetSingleTarget();
                var snipId = cells.GetSnipId(target);
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
                EnsureProject();
                var target = cells.GetSingleTarget();
                var snipId = cells.GetSnipId(target);
                var snip = context.State.Snips.FirstOrDefault(item => item.Id == snipId);
                if (snip == null) throw new InvalidOperationException("La cellule sélectionnée n'a aucune preuve Doctracker.");

                using (var dialog = new ReviewDialog(snip))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    context.Snips.SetReview(context.State, snip.Id, dialog.SelectedStatus,
                        dialog.ReviewComment, Environment.UserName);
                }

                var document = context.State.Documents.First(item => item.Id == snip.DocumentId);
                cells.AttachProof(target, snip, document);
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

            SelectDocument(document.Id);
            canvas.NavigateTo(context.Store.ResolveDocumentPath(document), snip);
            SetStatus(document.OriginalName + " — page " + snip.PageNumber + " — " + snip.Status);
            return true;
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.FromArgb(255, 122, 0)
            };
            var title = new Label
            {
                Text = "DOCTRACKER",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold),
                Location = new Point(12, 7),
                AutoSize = true
            };
            var subtitle = new Label
            {
                Text = "Preuve locale • OCR • Recherche • Traçabilité",
                ForeColor = Color.FromArgb(255, 244, 230),
                Location = new Point(14, 34),
                AutoSize = true
            };
            header.Controls.Add(title);
            header.Controls.Add(subtitle);
            return header;
        }

        private Control BuildSearchBar(out TextBox queryBox)
        {
            var bar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                Padding = new Padding(8, 5, 8, 5),
                BackColor = Color.White
            };
            queryBox = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };
            var searchButton = new Button
            {
                Text = "Rechercher",
                Dock = DockStyle.Right,
                Width = 92,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(18, 28, 45),
                ForeColor = Color.White,
                TabStop = false
            };
            searchButton.FlatAppearance.BorderSize = 0;
            searchButton.Click += (sender, args) => SearchFromPane();
            queryBox.KeyDown += (sender, args) =>
            {
                if (args.KeyCode == Keys.Enter)
                {
                    args.SuppressKeyPress = true;
                    SearchFromPane();
                }
            };
            bar.Controls.Add(queryBox);
            bar.Controls.Add(searchButton);
            return bar;
        }

        private Control BuildWorkflowBar(out Label label)
        {
            var bar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(10, 5, 5, 0),
                BackColor = Color.FromArgb(248, 249, 251)
            };
            label = new Label
            {
                Text = "Mode actif : aucun — choisissez un type dans le ruban",
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = Color.FromArgb(80, 88, 98)
            };
            bar.Controls.Add(label);
            return bar;
        }

        private void Documents_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var document = SelectedDocument;
                if (document != null)
                {
                    canvas.LoadDocument(context.Store.ResolveDocumentPath(document));
                    SetStatus(document.OriginalName + " — sélectionnez une zone ou recherchez dans les pièces.");
                }
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
                canvas.NavigateTo(context.Store.ResolveDocumentPath(result.Document), result.Candidate.PageNumber);
                SetStatus(result.Document.OriginalName + " — page " + result.Candidate.PageNumber +
                          " — score " + result.Candidate.Score.ToString("P0"));
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private DocumentRecord SelectedDocument => documents.SelectedItem as DocumentRecord;

        private void EnsureProject()
        {
            context.Ensure(application.ActiveWorkbook);
        }

        private void BindDocuments()
        {
            var selectedId = SelectedDocument?.Id;
            documents.SelectedIndexChanged -= Documents_SelectedIndexChanged;
            documents.DataSource = null;
            documents.DataSource = context.State.Documents.ToList();
            documents.DisplayMember = "OriginalName";
            documents.SelectedIndexChanged += Documents_SelectedIndexChanged;
            if (selectedId != null) SelectDocument(selectedId);
            if (documents.SelectedIndex < 0 && documents.Items.Count > 0) documents.SelectedIndex = 0;

            var total = context.State.Documents.Count;
            var indexed = context.State.Documents.Count(item => item.IndexedPages.Count > 0);
            ocrState.Text = total == 0
                ? "Aucune pièce"
                : total + " pièce(s) • " + indexed + " indexée(s) par OCR";
        }

        private void SelectDocument(string id)
        {
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
            if (range == null || range.Columns.Count != 1)
                throw new InvalidOperationException("Sélectionnez une seule colonne pour le " + purpose + ".");
            if (range.Rows.Count > 5000)
                throw new InvalidOperationException("Limitez la sélection de " + purpose + " à 5 000 lignes.");
        }

        private ExcelInterop.Range ResolveMatchingRange(string worksheetName, string address)
        {
            if (string.IsNullOrWhiteSpace(worksheetName) || string.IsNullOrWhiteSpace(address)) return null;
            var workbook = application.ActiveWorkbook;
            if (workbook == null) return null;
            var worksheet = workbook.Worksheets[worksheetName] as ExcelInterop.Worksheet;
            return worksheet == null ? null : worksheet.Range[address];
        }

        private void SetStatus(string message)
        {
            status.Text = message;
        }

        private void SetStatusThreadSafe(string message)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(SetStatus), message);
            }
            else
            {
                SetStatus(message);
            }
        }

        private void ShowError(Exception exception)
        {
            SetStatus("Erreur : " + exception.Message);
            MessageBox.Show(this, exception.Message, "Doctracker",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private sealed class SearchResultItem
        {
            public MatchCandidate Candidate { get; set; }
            public DocumentRecord Document { get; set; }
            public string Caption => Document.OriginalName + " — page " + Candidate.PageNumber +
                                     " — score " + Candidate.Score.ToString("P0");
        }
    }
}
