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
            view.Comment.Click += (s,e) => {
                if(context.IsBusy)return;
                activeSnipType=null;view.SetReadingMode(false);view.SetCommentMode(!canvas.CommentMode);
                Ribbon.DoctrackerRibbon.Instance?.Refresh();
                SetStatus(canvas.CommentMode ? "Dessinez le cadre du commentaire sur le document." : "Commentaire désactivé.");
            };
            view.DeleteSnip.Click += (s,e) => DeleteSnip(focusedSnipId);
            view.DeleteSnipMenu.Click += (s,e) => DeleteSnip(focusedSnipId);
            canvas.DeleteProofRequested += DeleteSnip;
            canvas.EditCommentRequested += EditDocumentComment;
            canvas.DeleteCommentRequested += DeleteDocumentComment;
            canvas.ProofSelected += id => {
                if(context.IsBusy)return; focusedSnipId=id;
                var snip=context.State.Snips.FirstOrDefault(s=>s.Id==id);if(snip==null)return;
                bindingProofs=true;
                try { view.Proofs.DataSource=new[]{new ProofItem {Id=id,Caption=SnipTheme.LabelFor(snip.SourceType??snip.Type)+" · "+snip.WorksheetName+"!"+snip.CellAddress}};view.ShowProofs(true); }
                finally {bindingProofs=false;}
            };
        }
        private string CommentText(string text)
        {
            using(var dialog=new Form {Text="Commentaire sur le document",ClientSize=new Size(480,250),MinimumSize=new Size(360,220),
                StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10),Padding=new Padding(12),MinimizeBox=false,MaximizeBox=false})
            {
                var input=new TextBox {Text=text,Multiline=true,AcceptsReturn=true,ScrollBars=ScrollBars.Vertical,Dock=DockStyle.Fill,MaxLength=2000,ForeColor=Color.Red};
                var footer=new FlowLayoutPanel {Dock=DockStyle.Bottom,AutoSize=true,FlowDirection=FlowDirection.RightToLeft};
                var ok=new Button {Text="Enregistrer",AutoSize=true,DialogResult=DialogResult.OK};
                var cancel=new Button {Text="Annuler",AutoSize=true,DialogResult=DialogResult.Cancel};
                footer.Controls.Add(ok);footer.Controls.Add(cancel);dialog.Controls.Add(input);dialog.Controls.Add(footer);dialog.CancelButton=cancel;
                dialog.Shown+=(s,e)=>input.Focus();
                return dialog.ShowDialog(this)==DialogResult.OK ? input.Text : null;
            }
        }
        private void CreateDocumentComment()
        {
            try
            {
                EnsureProject();var doc=SelectedDocument;if(doc==null||!canvas.HasSelection)return;
                var text=CommentText("");if(text==null)return;
                var zone=canvas.FitComment(canvas.GetNormalizedSelection(),text);
                new DocumentCommentService(context.Store).Save(context.State,doc.Id,canvas.CurrentPageNumber,
                    new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),text);
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
                var text=CommentText(comment.Text);if(text==null)return;
                var zone=canvas.FitComment(new RectangleF((float)comment.X,(float)comment.Y,(float)comment.Width,(float)comment.Height),text);
                new DocumentCommentService(context.Store).Save(context.State,doc.Id,comment.PageNumber,
                    new NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),text,id);
                UpdateDocumentProofs();context.MarkWorkbookDirty();SetStatus("Commentaire enregistré.");
            }
            catch(Exception exception){ShowError(exception);}
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
