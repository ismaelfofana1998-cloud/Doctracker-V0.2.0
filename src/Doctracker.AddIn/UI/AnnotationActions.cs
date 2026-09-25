using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Doctracker.AddIn.Excel;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed partial class DoctrackerPaneControl
    {
        private void WireAnnotationActions()
        {
            view.DeleteSnip.Click += (s,e) => DeleteSnip(focusedSnipId);
            canvas.DeleteProofRequested += DeleteSnip;
            canvas.EditCommentRequested += EditDocumentComment;
            canvas.CommentGeometryChanged += ResizeDocumentComment;
            canvas.ProofGeometryChanged += ResizeSnip;
            canvas.DeleteCommentRequested += DeleteDocumentComment;
            canvas.ProofSelected += id => {
                if(context.IsBusy)return; focusedSnipId=id;
                var snip=context.State.Snips.FirstOrDefault(s=>s.Id==id);if(snip==null)return;
                bindingProofs=true;
                try { view.Proofs.DataSource=new[]{new ProofItem {Id=id,Caption=SnipTheme.LabelFor(snip.SourceType??snip.Type)+" · "+snip.WorksheetName+"!"+snip.CellAddress}};view.ShowProofs(true); }
                finally {bindingProofs=false;}
                NavigateToLinkedCell(snip);
            };
        }
        private void ToggleComment()
        {
            if(context.IsBusy)return;
            var enabled=!canvas.CommentMode;
            activeSnipType=enabled?(SnipType?)null:SnipType.Text;view.SetCommentMode(enabled);
            if(!enabled)view.SetMode(SnipType.Text);
            Ribbon.DoctrackerRibbon.Instance?.Refresh();
            SetStatus(canvas.CommentMode ? "Commentaire" : "Prêt");
        }
        private void NavigateToLinkedCell(SnipRecord snip)
        {
            var events=application.EnableEvents;
            try
            {
                EnsureActiveWorkbook();
                ExcelProofLinks.Invalidate(workbook);
                // Validate the live marker: recorded addresses can be stale after row moves.
                var current=application.Selection as ExcelInterop.Range;
                if(current!=null && current.Cells.CountLarge==1 && cells.GetSnipIds(current).Contains(snip.Id))return;
                ExcelInterop.Range target=null;
                foreach(ExcelInterop.Worksheet sheet in workbook.Worksheets)
                {
                    if(sheet.Name!=snip.WorksheetName)continue;
                    var recorded=sheet.Range[snip.CellAddress];
                    if(recorded.Cells.CountLarge==1 && cells.GetSnipIds(recorded).Contains(snip.Id))target=recorded;
                }
                if(target==null)
                {
                    target=cells.LinkedCells(workbook).FirstOrDefault(cell=>cells.GetSnipIds(cell).Contains(snip.Id));
                }
                if(target==null)
                {
                    if(current==null || current.Cells.CountLarge!=1)
                        throw new InvalidOperationException("Le lien Excel a disparu. Sélectionnez la cellule de destination, puis cliquez de nouveau sur cette preuve pour la relier.");
                    ExcelCellGateway.ValidateWritable(current);
                    if(MessageBox.Show(this,"Le lien Excel a disparu ou sa cellule a été supprimée. Relier cette preuve à la cellule sélectionnée "+
                        current.Worksheet.Name+"!"+current.Address[false,false]+" ? Sa valeur sera conservée.","Relier la preuve",
                        MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;
                    var snapshot=cells.Snapshot(current);application.EnableEvents=false;
                    try
                    {
                        cells.AttachProof(current,snip,context.State.Documents.First(d=>d.Id==snip.DocumentId));
                        context.CaptureCellLinks();workbook.Saved=false;
                    }
                    catch {snapshot.Restore();throw;}
                    target=current;
                }
                if(target.Worksheet.Visible!=ExcelInterop.XlSheetVisibility.xlSheetVisible)
                    throw new InvalidOperationException("La preuve est liée à une feuille masquée : "+target.Worksheet.Name+". Affichez cette feuille pour y accéder.");
                application.EnableEvents=false;application.Goto(target,true);
                SetStatus("Cellule liée : "+target.Worksheet.Name+"!"+target.Address[false,false]);
            }
            catch(Exception exception){ShowError(exception);}
            finally{application.EnableEvents=events;}
        }
        private DocumentComment CommentProperties(DocumentComment original)
        {
            using(var dialog=new Form {Text="Commentaire : texte et dimensions",ClientSize=new Size(520,330),MinimumSize=new Size(460,300),
                StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10),Padding=new Padding(12),MinimizeBox=false,MaximizeBox=false})
            {
                var input=new TextBox {Text=original.Text,Multiline=true,AcceptsReturn=true,ScrollBars=ScrollBars.Vertical,Dock=DockStyle.Fill,MaxLength=2000,ForeColor=Color.Red};
                var settings=new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true,WrapContents=true,Padding=new Padding(0,0,0,8)};
                var fontSize=new NumericUpDown {Minimum=6,Maximum=72,Value=(decimal)Math.Max(6,Math.Min(72,original.FontSize)),Width=64};
                var width=new NumericUpDown {Minimum=1,Maximum=(decimal)Math.Max(1,Math.Floor((1-original.X)*100)),DecimalPlaces=1,Width=70};
                var height=new NumericUpDown {Minimum=1,Maximum=(decimal)Math.Max(1,Math.Floor((1-original.Y)*100)),DecimalPlaces=1,Width=70};
                width.Value=Math.Min(width.Maximum,Math.Max(1,(decimal)original.Width*100));height.Value=Math.Min(height.Maximum,Math.Max(1,(decimal)original.Height*100));
                settings.Controls.Add(new Label {Text="Texte (pt)",AutoSize=true});settings.Controls.Add(fontSize);
                settings.Controls.Add(new Label {Text="Largeur (%)",AutoSize=true});settings.Controls.Add(width);
                settings.Controls.Add(new Label {Text="Hauteur (%)",AutoSize=true});settings.Controls.Add(height);
                var footer=new FlowLayoutPanel {Dock=DockStyle.Bottom,AutoSize=true,FlowDirection=FlowDirection.RightToLeft};
                var ok=new Button {Text="Enregistrer",AutoSize=true,DialogResult=DialogResult.OK};
                var cancel=new Button {Text="Annuler",AutoSize=true,DialogResult=DialogResult.Cancel};
                footer.Controls.Add(ok);footer.Controls.Add(cancel);dialog.Controls.Add(input);dialog.Controls.Add(settings);dialog.Controls.Add(footer);dialog.CancelButton=cancel;
                Font previewFont=null;
                Action updateFont=()=>{var old=previewFont;previewFont=new Font("Segoe UI",(float)fontSize.Value);input.Font=previewFont;old?.Dispose();};
                fontSize.ValueChanged+=(s,e)=>updateFont();updateFont();dialog.Shown+=(s,e)=>input.Focus();
                dialog.FormClosing+=(s,e)=>{
                    if(dialog.DialogResult!=DialogResult.OK)return;
                    try
                    {
                        if(string.IsNullOrWhiteSpace(input.Text))throw new InvalidOperationException("Saisissez le texte du commentaire.");
                        var fitted=canvas.FitComment(new RectangleF((float)original.X,(float)original.Y,(float)width.Value/100,(float)height.Value/100),input.Text,(double)fontSize.Value);
                        // FitComment expands the height as needed; the saved frame includes all text.
                        height.Value=Math.Min(height.Maximum,Math.Max(height.Minimum,(decimal)fitted.Height*100));
                    }
                    catch(Exception exception){e.Cancel=true;MessageBox.Show(dialog,exception.Message,"Commentaire",MessageBoxButtons.OK,MessageBoxIcon.Information);}
                };
                try
                {
                    if(dialog.ShowDialog(this)!=DialogResult.OK)return null;
                    return new DocumentComment {Id=original.Id,PageNumber=original.PageNumber,X=original.X,Y=original.Y,
                        Width=(double)width.Value/100,Height=(double)height.Value/100,FontSize=(double)fontSize.Value,Text=input.Text};
                }
                finally{input.Font=dialog.Font;previewFont?.Dispose();}
            }
        }
        private void CreateDocumentComment()
        {
            try
            {
                EnsureProject();var doc=SelectedDocument;if(doc==null||!canvas.HasSelection)return;
                var selected=canvas.GetNormalizedSelection();
                var edited=CommentProperties(new DocumentComment {X=selected.X,Y=selected.Y,Width=selected.Width,Height=selected.Height});if(edited==null)return;
                var zone=canvas.FitComment(new RectangleF((float)edited.X,(float)edited.Y,(float)edited.Width,(float)edited.Height),edited.Text,edited.FontSize);
                new DocumentCommentService(context.Store).Save(context.State,doc.Id,canvas.CurrentPageNumber,
                    new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),edited.Text,fontSize:edited.FontSize);
                canvas.ClearSelection();UpdateDocumentProofs();context.MarkWorkbookDirty();SetStatus("Commentaire ajouté. Clic droit sur le cadre pour le modifier ou le supprimer.");
            }
            catch(Exception exception){ShowError(exception);}
        }
        private void EditDocumentComment(string id)
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var doc=SelectedDocument;var comment=doc?.Comments.FirstOrDefault(c=>c.Id==id);if(comment==null)return;
                var edited=CommentProperties(comment);if(edited==null)return;
                var zone=canvas.FitComment(new RectangleF((float)edited.X,(float)edited.Y,(float)edited.Width,(float)edited.Height),edited.Text,edited.FontSize);
                new DocumentCommentService(context.Store).Save(context.State,doc.Id,comment.PageNumber,
                    new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),edited.Text,id,edited.FontSize);
                UpdateDocumentProofs();context.MarkWorkbookDirty();SetStatus("Commentaire enregistré.");
            }
            catch(Exception exception){ShowError(exception);}
        }
        private void ResizeDocumentComment(string id,RectangleF zone)
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var doc=SelectedDocument;var comment=doc?.Comments.FirstOrDefault(c=>c.Id==id);if(comment==null)return;
                var fitted=canvas.FitComment(zone,comment.Text,comment.FontSize);
                if(fitted.Height>zone.Height+.002)throw new InvalidOperationException("Le cadre est trop petit pour ce texte. Agrandissez-le ou réduisez la police par double-clic. Le cadre précédent est conservé.");
                new DocumentCommentService(context.Store).Save(context.State,doc.Id,comment.PageNumber,
                    new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),comment.Text,id,comment.FontSize);
                UpdateDocumentProofs();context.MarkWorkbookDirty();SetStatus("Commentaire déplacé / redimensionné. Double-clic pour modifier le texte et la police.");
            }
            catch(Exception exception){ShowError(exception);canvas.Invalidate();}
        }
        private void DeleteDocumentComment(string id)
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();if(SelectedDocument==null)return;
                if(MessageBox.Show(this,"Supprimer ce commentaire ?","Commentaire",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                new DocumentCommentService(context.Store).Delete(context.State,SelectedDocument.Id,id);
                UpdateDocumentProofs();context.MarkWorkbookDirty();SetStatus("Commentaire supprimé.");
            }
            catch(Exception exception){ShowError(exception);}
        }
        private async void ResizeSnip(string id,RectangleF zone)
        {
            if(context.IsBusy)return;
            var snapshots=new List<ExcelCellGateway.CellSnapshot>();var events=application.EnableEvents;
            SnipRecord snip=null;DocumentRecord doc=null;
            try
            {
                EnsureProject();snip=context.State.Snips.FirstOrDefault(s=>s.Id==id);
                doc=snip==null?null:context.State.Documents.FirstOrDefault(d=>d.Id==snip.DocumentId);
                if(doc==null)return;
                var targets=cells.LinkedCells(workbook).Where(cell=>cells.GetSnipIds(cell).Contains(id)).ToList();
                foreach(var target in targets)ExcelCellGateway.ValidateWritable(target);
                var preserve=snip.Type==SnipType.Validation || snip.Type==SnipType.Exception;
                // A changed cell/formula is not silently overwritten by moving its proof.
                if(!preserve && targets.Any(target=>CellChangedSinceSnip(target,snip)) &&
                    MessageBox.Show(this,"La valeur ou la formule liée a été modifiée. La remplacer par la nouvelle extraction ?","Modifier le snip",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                BeginOperation();SetStatus(preserve?"Mise à jour de la zone…":"Lecture de la nouvelle zone…");
                var raw=snip.RawText;
                if(!preserve)
                {
                    var recognized=ExtractIndexedSelection(doc,snip.PageNumber,zone);
                    if(recognized==null)
                    {
                        var path=canvas.CurrentPath;
                        var page=await System.Threading.Tasks.Task.Run(()=>DocumentIndexer.ReadNativePage(path,snip.PageNumber));
                        recognized=ExtractPageSelection(page,zone);
                    }
                    if(recognized==null)
                    {
                        SetStatus("Reconnaissance de la nouvelle zone…");
                        using(var crop=canvas.CropRegion(zone))recognized=await System.Threading.Tasks.Task.Run(()=>ocr.Recognize(crop));
                    }
                    raw=snip.Type==SnipType.Sum?SumSourceText(recognized):recognized.Text;
                }
                operation.Token.ThrowIfCancellationRequested();EnsureActiveWorkbook();
                foreach(var target in targets)snapshots.Add(cells.Snapshot(target));
                application.EnableEvents=false;
                context.Snips.UpdateGeometry(context.State,id,new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),raw,Environment.UserName,updated=>{
                    foreach(var target in targets)
                    {
                        if(preserve){cells.AttachProof(target,updated,doc,true);continue;}
                        var linked=cells.GetSnipIds(target).Select(key=>context.State.Snips.FirstOrDefault(s=>s.Id==key)).Where(s=>s!=null).ToList();
                        cells.WriteSnip(target,updated,doc,true);
                        if(updated.Type==SnipType.Sum && linked.All(s=>s.Type==SnipType.Sum))
                            target.Value2=(double)linked.Sum(s=>decimal.Parse(s.ExtractedValue,System.Globalization.CultureInfo.InvariantCulture));
                    }
                });
                snapshots.Clear();SetStatus("Snip mis à jour · "+targets.Count+" cellule(s) actualisée(s).");
            }
            catch(OperationCanceledException){SetStatus("Modification du snip annulée.");}
            catch(Exception failure)
            {
                var errors=new List<string>();foreach(var snapshot in snapshots)try{snapshot.Restore();}catch(Exception rollback){errors.Add(rollback.Message);}
                ShowError(errors.Count==0?failure:new InvalidOperationException(failure.Message+" — Restauration Excel incomplète : "+string.Join(" ; ",errors)));
            }
            finally
            {
                application.EnableEvents=events;EndOperation();UpdateDocumentProofs();
                if(snip!=null && doc!=null && canvas.CurrentPath!=null)canvas.NavigateTo(canvas.CurrentPath,snip);
            }
        }
        private bool CellChangedSinceSnip(ExcelInterop.Range target,SnipRecord snip)
        {
            if(Equals(target.HasFormula,true))return true;
            if(snip.Type==SnipType.Sum || snip.Type==SnipType.Number)
            {
                var expected=decimal.Parse(snip.ExtractedValue,System.Globalization.CultureInfo.InvariantCulture);
                if(snip.Type==SnipType.Sum)
                {
                    var linked=cells.GetSnipIds(target).Select(id=>context.State.Snips.FirstOrDefault(s=>s.Id==id)).ToList();
                    if(linked.Any(s=>s==null || s.Type!=SnipType.Sum))return true;
                    expected=linked.Sum(s=>decimal.Parse(s.ExtractedValue,System.Globalization.CultureInfo.InvariantCulture));
                }
                decimal actual;
                return !decimal.TryParse(Convert.ToString(target.Value2,System.Globalization.CultureInfo.InvariantCulture),System.Globalization.NumberStyles.Any,System.Globalization.CultureInfo.InvariantCulture,out actual) || actual!=expected;
            }
            return ExcelCellGateway.QueryText(target)!=snip.ExtractedValue;
        }

        private void DeleteSelectionSnips()
        {
            if(context.IsBusy)return;
            var snapshots=new List<ExcelCellGateway.CellSnapshot>();var events=application.EnableEvents;
            try
            {
                EnsureProject();var selection=cells.GetSelection();
                var targets=new List<ExcelInterop.Range>();
                // Iterate proof cells, never all 1,048,576 cells of a selected column.
                foreach(var cell in cells.LinkedCells(workbook))
                    if(cell.Worksheet.CodeName==selection.Worksheet.CodeName && application.Intersect(cell,selection)!=null)
                    {ExcelCellGateway.ValidateWritable(cell);targets.Add(cell);}
                if(targets.Count==0){SetStatus("Aucun snip lié dans la plage sélectionnée.");return;}
                if(MessageBox.Show(this,"Retirer les snips de "+targets.Count+" cellule(s) dans "+selection.Address[false,false]+
                    " ? Les valeurs et formules seront conservées. Les liens hors de cette plage seront conservés.",
                    "Supprimer les snips de la plage",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;
                context.CaptureCellLinks();
                var selected=targets.Select(cell=>new CellLinkRecord {WorksheetName=cell.Worksheet.Name,
                    WorksheetCodeName=cell.Worksheet.CodeName,CellAddress=cell.Address[false,false,ExcelInterop.XlReferenceStyle.xlA1],
                    SnipIds=cells.GetSnipIds(cell).ToList()}).ToList();
                foreach(var target in targets)snapshots.Add(cells.Snapshot(target));
                application.EnableEvents=false;
                foreach(var target in targets)cells.RemoveProof(target);
                var deleted=context.Snips.DeleteCellLinks(context.State,selected,Environment.UserName);
                snapshots.Clear();focusedSnipId=null;view.ShowProofs(false);canvas.ClearSelection();UpdateDocumentProofs();context.MarkWorkbookDirty();
                SetStatus(targets.Count+" cellule(s) déliée(s), "+deleted+" snip(s) supprimé(s). Valeurs conservées.");
            }
            catch(Exception exception)
            {
                var failures=new List<string>();
                foreach(var snapshot in snapshots)try{snapshot.Restore();}catch(Exception rollback){failures.Add(rollback.Message);}
                ShowError(failures.Count==0?exception:new InvalidOperationException(exception.Message+"\nRestauration de cellules incomplète : "+string.Join(" ; ",failures)));
            }
            finally {application.EnableEvents=events;}
        }

        private void DeleteSnip(string id)
        {
            if(context.IsBusy)return;
            var snapshots=new List<ExcelCellGateway.CellSnapshot>();var events=application.EnableEvents;
            try
            {
                EnsureProject();
                if(string.IsNullOrEmpty(id))id=ChooseSnipId(cells.GetSingleTarget());
                if(string.IsNullOrEmpty(id)||!context.State.Snips.Any(s=>s.Id==id))throw new InvalidOperationException("Sélectionnez une cellule liée ou faites un clic droit sur le snip du document.");
                if(MessageBox.Show(this,"Supprimer ce snip et tous ses liens dans ce classeur ? Les valeurs et formules Excel seront conservées, y compris les totaux de sommes : vérifiez-les si nécessaire.","Supprimer le snip",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;
                var targets=new List<ExcelInterop.Range>();
                foreach(var cell in cells.LinkedCells(workbook))
                    if(cells.GetSnipIds(cell).Contains(id)){ExcelCellGateway.ValidateWritable(cell);targets.Add(cell);}
                foreach(var target in targets)snapshots.Add(cells.Snapshot(target));
                application.EnableEvents=false;
                foreach(var target in targets)cells.DetachProof(target,id,context.State);
                context.Snips.Delete(context.State,id,Environment.UserName);
                // No further fallible Excel writes after the durable metadata commit.
                snapshots.Clear();focusedSnipId=null;view.ShowProofs(false);canvas.ClearSelection();UpdateDocumentProofs();context.MarkWorkbookDirty();
                SetStatus("Snip supprimé de "+targets.Count+" cellule(s). Les valeurs Excel sont conservées.");
            }
            catch(Exception exception)
            {
                var failures=new List<string>();
                foreach(var snapshot in snapshots)try{snapshot.Restore();}catch(Exception rollback){failures.Add(rollback.Message);}
                ShowError(failures.Count==0?exception:new InvalidOperationException(exception.Message+"\nRestauration de cellules incomplète : "+string.Join(" ; ",failures)));
            }
            finally {application.EnableEvents=events;}
        }
    }
}
