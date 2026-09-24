using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Doctracker.AddIn.Excel;
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
            activeSnipType=null;view.SetCommentMode(!canvas.CommentMode);
            Ribbon.DoctrackerRibbon.Instance?.Refresh();
            SetStatus(canvas.CommentMode ? "Dessinez le cadre du commentaire sur le document." : "Commentaire désactivé.");
        }
        private void NavigateToLinkedCell(SnipRecord snip)
        {
            var events=application.EnableEvents;
            try
            {
                EnsureActiveWorkbook();
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
                    foreach(ExcelInterop.Worksheet sheet in workbook.Worksheets)
                    {
                        ExcelInterop.Range linked;
                        try{linked=sheet.Cells.SpecialCells(ExcelInterop.XlCellType.xlCellTypeComments);}
                        catch(System.Runtime.InteropServices.COMException){continue;}
                        foreach(ExcelInterop.Range cell in linked.Cells)
                            if(cells.GetSnipIds(cell).Contains(snip.Id)){target=cell;break;}
                        if(target!=null)break;
                    }
                }
                if(target==null)throw new InvalidOperationException("La cellule liée est introuvable. Utilisez Récupération > Réparer les liens dans le ruban.");
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
                foreach(ExcelInterop.Worksheet sheet in workbook.Worksheets)
                {
                    ExcelInterop.Range comments;
                    try {comments=sheet.Cells.SpecialCells(ExcelInterop.XlCellType.xlCellTypeComments);}
                    catch(System.Runtime.InteropServices.COMException){continue;}
                    foreach(ExcelInterop.Range cell in comments.Cells)
                        if(cells.GetSnipIds(cell).Contains(id)){ExcelCellGateway.ValidateWritable(cell);targets.Add(cell);}
                }
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
