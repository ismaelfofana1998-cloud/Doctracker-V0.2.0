using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed partial class DoctrackerPaneControl
    {
        private bool bindingCategories;
        private void WireProjectActions()
        {
            view.ImportFolder.Click+=(s,e)=>ImportFolderAsync();
            view.Categorize.Click+=(s,e)=>ChangeCategory();
            view.CrossReference.Click+=(s,e)=>AssignReference();
            view.Backup.Click+=(s,e)=>BackupAsync();
            view.Restore.Click+=(s,e)=>RestoreBackup();
            view.RepairLinks.Click+=(s,e)=>RepairLinks();
            view.ExportPdf.Click+=(s,e)=>ExportPdfAsync();
            view.SharedStorage.Click+=(s,e)=>SelectStorage();
            view.RecoveryFolder.Click+=(s,e)=>System.Diagnostics.Process.Start("explorer.exe",WorkbookProjectContext.CacheRoot);
            view.Categories.SelectedIndexChanged+=(s,e)=>{if(!bindingCategories && !context.IsBusy)BindDocuments();};
        }
        private static string Prompt(IWin32Window owner,string title,string label,string value="")
        {
            using(var dialog=new Form {Text=title,Width=530,Height=175,StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10),MinimizeBox=false,MaximizeBox=false})
            {
                var input=new TextBox {Text=value,Dock=DockStyle.Top};var caption=new Label {Text=label,Dock=DockStyle.Top,AutoSize=true};
                var ok=new Button {Text="Valider",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK,AutoSize=true};
                dialog.Padding=new Padding(12);dialog.Controls.Add(input);dialog.Controls.Add(caption);dialog.Controls.Add(ok);dialog.AcceptButton=ok;
                return dialog.ShowDialog(owner)==DialogResult.OK?input.Text:null;
            }
        }
        private void RefreshCategories()
        {
            var selected=view.Categories.SelectedItem as string;
            bindingCategories=true;
            try
            {
                var values=new[]{"Toutes les catégories"}.Concat(context.State.Documents.SelectMany(d=>d.Categories).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(c=>c)).ToList();
                view.Categories.DataSource=values;view.Categories.SelectedItem=values.Contains(selected)?selected:values[0];
            }
            finally{bindingCategories=false;}
        }
        private IEnumerable<DocumentRecord> VisibleDocuments()
        {
            var category=view.Categories.SelectedItem as string;
            return string.IsNullOrEmpty(category)||category=="Toutes les catégories"?context.State.Documents:context.State.Documents.Where(d=>d.Categories.Contains(category));
        }
        private ProjectState SearchScope() => new ProjectState {Documents=VisibleDocuments().ToList()};
        private void ChangeCategory()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var doc=SelectedDocument;if(doc==null)return;
                var value=Prompt(this,"Catégories", "Catégories séparées par ; (ex. Client 1 ; Factures)",string.Join(" ; ",doc.Categories));if(value==null)return;
                var old=doc.Categories;doc.Categories=value.Split(';').Select(c=>c.Trim()).Where(c=>c.Length>0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                try{context.Store.Save(context.State);}catch{doc.Categories=old;throw;}
                RefreshCategories();BindDocuments();context.MarkWorkbookDirty();SetStatus("Catégories enregistrées.");
            }
            catch(Exception exception){ShowError(exception);}
        }
        private async void ImportFolderAsync()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();using(var picker=new FolderBrowserDialog {Description="Importer un dossier et ses sous-dossiers",ShowNewFolderButton=false})
                {
                    if(picker.ShowDialog(this)!=DialogResult.OK)return;
                    var category=Prompt(this,"Importer un dossier","Catégorie principale (les sous-dossiers sont conservés)",Path.GetFileName(picker.SelectedPath));if(category==null)return;
                    BeginOperation();var root=picker.SelectedPath;var errors=new List<string>();
                    var count=await Task.Run(()=>
                    {
                        var imported=0;
                        foreach(var file in EnumerateDocuments(root,errors))
                        {
                            operation.Token.ThrowIfCancellationRequested();
                            var relative=Path.GetDirectoryName(file).Substring(root.TrimEnd(Path.DirectorySeparatorChar).Length).Trim(Path.DirectorySeparatorChar);
                            try {context.Importer.Import(context.State,file,Environment.UserName,string.IsNullOrWhiteSpace(relative)?category:category+" / "+relative.Replace(Path.DirectorySeparatorChar,'/'),false);imported++;}
                            catch(Exception exception){errors.Add(Path.GetFileName(file)+" : "+exception.Message);}
                            if(imported%25==0){context.Store.Save(context.State);SetStatusThreadSafe(imported+" pièces importées…");}
                        }
                        context.Store.Save(context.State);return imported;
                    });
                    RefreshCategories();BindDocuments();errors.AddRange(await IndexMissingAsync());BindDocuments();
                    SetStatus(count+" pièces importées. "+errors.Count+" erreur(s).");if(errors.Count>0)MessageBox.Show(this,string.Join("\n",errors.Take(30)),"Import : pièces à vérifier");
                }
            }
            catch(OperationCanceledException){context.Store.Save(context.State);SetStatus("Import arrêté. Les pièces déjà importées sont conservées.");}
            catch(Exception exception){ShowError(exception);}
            finally{EndOperation();}
        }
        private static IEnumerable<string> EnumerateDocuments(string root,List<string> errors)
        {
            var queue=new Stack<string>();queue.Push(root);
            while(queue.Count>0)
            {
                var folder=queue.Pop();string[] files=null;
                try{files=Directory.GetFiles(folder);foreach(var child in Directory.GetDirectories(folder))if((File.GetAttributes(child)&FileAttributes.ReparsePoint)==0)queue.Push(child);}
                catch(Exception ex) when(ex is IOException||ex is UnauthorizedAccessException){errors.Add(folder+" : "+ex.Message);}
                foreach(var file in files??new string[0])if(DocumentImporter.IsSupported(file))yield return file;
            }
        }
        private void AssignReference()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var doc=SelectedDocument;if(doc==null)return;
                var reference=Prompt(this,"Xref du document","Référence du test (ex. DAC B 30 040)",doc.TestReference);if(string.IsNullOrWhiteSpace(reference))return;
                using(var dialog=new Form {Text="Numéro Xref disponible",Width=400,Height=150,StartPosition=FormStartPosition.CenterParent})
                {
                    var list=new ComboBox {Dock=DockStyle.Top,DropDownStyle=ComboBoxStyle.DropDownList,DataSource=CrossReferences.Available(context.State,reference).Take(500).ToList()};
                    list.Format+=(s,e)=>e.Value=((int)e.ListItem).ToString("D2");list.FormattingEnabled=true;
                    var ok=new Button {Text="Attribuer la Xref",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK};dialog.Controls.Add(list);dialog.Controls.Add(ok);
                    if(dialog.ShowDialog(this)!=DialogResult.OK)return;
                    var oldRef=doc.TestReference;var oldNumber=doc.ReferenceNumber;
                    CrossReferences.Assign(context.State,doc,reference,(int)list.SelectedItem);
                    try{context.Store.Save(context.State);}catch{doc.TestReference=oldRef;doc.ReferenceNumber=oldNumber;context.State.XrefReservations.RemoveAt(context.State.XrefReservations.Count-1);throw;}
                    documents.Refresh();BindDocuments();context.MarkWorkbookDirty();SetStatus("Xref : "+doc.DisplayName+". Le nom de l'export reprend cette référence.");
                }
            }
            catch(Exception exception){ShowError(exception);}
        }
        private async void BackupAsync()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();using(var dialog=new SaveFileDialog {Filter="Sauvegarde Doctracker|*.dtpack",DefaultExt="dtpack",AddExtension=true,FileName="Doctracker-"+DateTime.Today.ToString("yyyyMMdd")+".dtpack"})
                {if(dialog.ShowDialog(this)!=DialogResult.OK)return;context.CaptureCellLinks();BeginOperation();await Task.Run(()=>RecoveryArchive.Export(context.Store,context.State,dialog.FileName,operation.Token));SetStatus("Sauvegarde complète créée : pièces, catégories, index et liens.");}
            }
            catch(OperationCanceledException){SetStatus("Sauvegarde annulée.");}catch(Exception ex){ShowError(ex);}finally{EndOperation();}
        }
        private void RestoreBackup()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();using(var dialog=new OpenFileDialog {Filter="Sauvegarde Doctracker|*.dtpack;*.xml"})
                {
                    if(dialog.ShowDialog(this)!=DialogResult.OK)return;
                    if(MessageBox.Show(this,"Restaurer ce dossier dans le classeur actif ? Les valeurs Excel restent inchangées. L'ancien dossier reste dans les sauvegardes locales.","Restaurer",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                    if(Path.GetExtension(dialog.FileName).Equals(".xml",StringComparison.OrdinalIgnoreCase))context.RestoreMetadata(dialog.FileName);else context.RestoreArchive(dialog.FileName);canvas.ClearDocument();RefreshCategories();BindDocuments();RepairLinks();SetStatus("Dossier restauré. Enregistrez le classeur pour y intégrer les pièces.");
                }
            }
            catch(Exception ex){ShowError(ex);}
        }
        private void RepairLinks()
        {
            if(context.IsBusy)return;
            var snapshots=new List<Excel.ExcelCellGateway.CellSnapshot>();var events=application.EnableEvents;
            try
            {
                EnsureProject();if(MessageBox.Show(this,"Réattacher les preuves aux cellules enregistrées dans les métadonnées, sans modifier leurs valeurs ? Les feuilles manquantes seront signalées.","Récupération des liens",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                application.EnableEvents=false;var repaired=0;var missing=new List<string>();
                var links=context.State.CellLinks.ToList();
                var recorded=new HashSet<string>(links.SelectMany(link=>link.SnipIds));
                links.AddRange(context.State.Snips.Where(s=>!recorded.Contains(s.Id)).GroupBy(s=>new{s.WorksheetName,s.CellAddress}).Select(g=>new CellLinkRecord {WorksheetName=g.Key.WorksheetName,CellAddress=g.Key.CellAddress,SnipIds=g.Select(s=>s.Id).ToList()}));
                foreach(var group in links)
                {
                    ExcelInterop.Worksheet sheet=null;foreach(ExcelInterop.Worksheet item in workbook.Worksheets)if(item.Name==group.WorksheetName){sheet=item;break;}
                    if(sheet==null){missing.Add(group.WorksheetName);continue;}
                    var target=sheet.Range[group.CellAddress];if(target.Cells.CountLarge!=1)throw new InvalidDataException("Adresse de preuve invalide.");
                    Excel.ExcelCellGateway.ValidateWritable(target);snapshots.Add(cells.Snapshot(target));
                    foreach(var snip in context.State.Snips.Where(s=>group.SnipIds.Contains(s.Id)))cells.AttachProof(target,snip,context.State.Documents.First(d=>d.Id==snip.DocumentId));repaired++;
                }
                SetStatus(repaired+" cellule(s) réparée(s). Feuilles manquantes : "+string.Join(", ",missing.Distinct()));
                if(missing.Count>0)MessageBox.Show(this,"Créez ou renommez ces feuilles puis relancez la réparation : "+string.Join(", ",missing.Distinct()),"Liens non restaurés");
            }
            catch(Exception ex){foreach(var snapshot in snapshots)try{snapshot.Restore();}catch{}ShowError(ex);}
            finally{application.EnableEvents=events;}
        }
        private async void ExportPdfAsync()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var docs=VisibleDocuments().ToList();if(docs.Count==0)return;
                using(var dialog=new FolderBrowserDialog {Description="Exporter les documents de la catégorie active en PDF annotés"})
                {
                    if(dialog.ShowDialog(this)!=DialogResult.OK)return;context.CaptureCellLinks();BeginOperation();var folder=dialog.SelectedPath;
                    var errors=await Task.Run(()=>
                    {
                        var failures=new List<string>();
                        foreach(var doc in docs)
                        {
                            operation.Token.ThrowIfCancellationRequested();
                            var name=CrossReferences.SafeFileName(Path.GetFileNameWithoutExtension(doc.DisplayName))+".pdf";
                            var path=Path.Combine(folder,name);if(File.Exists(path))path=Path.Combine(folder,Path.GetFileNameWithoutExtension(name)+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".pdf");
                            try{AnnotatedPdfExporter.Export(context.Store.ResolveDocumentPath(doc),doc,context.State.Snips.Where(s=>s.DocumentId==doc.Id).ToArray(),path,operation.Token);}
                            catch(OperationCanceledException){throw;}catch(Exception ex){failures.Add(doc.DisplayName+" : "+ex.Message);}
                            SetStatusThreadSafe("Export : "+doc.DisplayName);
                        }
                        return failures;
                    });
                    SetStatus((docs.Count-errors.Count)+" PDF annotés exportés. Les originaux sont conservés.");if(errors.Count>0)MessageBox.Show(this,string.Join("\n",errors),"Exports à vérifier");
                }
            }
            catch(OperationCanceledException){SetStatus("Export arrêté. Les PDF terminés sont conservés.");}catch(Exception ex){ShowError(ex);}finally{EndOperation();}
        }
        private async void SelectStorage()
        {
            if(context.IsBusy)return;
            try
            {
                EnsureProject();var response=MessageBox.Show(this,"Oui : documents dans un emplacement réseau partagé (Excel léger).\nNon : pièces intégrées dans Excel (jusqu'à 256 Mo).\nLe partage réseau exige les droits d'accès pour chaque collègue.","Stockage des pièces",MessageBoxButtons.YesNoCancel);
                if(response==DialogResult.Cancel)return;var old=context.State.SharedVaultPath;
                if(response==DialogResult.Yes)
                {
                    using(var dialog=new FolderBrowserDialog {Description="Emplacement réseau commun, chemin UNC \\\\serveur\\partage"})
                    {
                        if(dialog.ShowDialog(this)!=DialogResult.OK)return;
                        if(!dialog.SelectedPath.StartsWith(@"\\"))throw new InvalidOperationException("Choisissez un chemin réseau UNC (\\\\serveur\\partage), identique pour tous les utilisateurs. Un dossier OneDrive personnel ou une lettre de lecteur locale n'est pas portable.");
                        var destination=dialog.SelectedPath;BeginOperation();
                        await Task.Run(()=>SharedVault.Publish(destination,context.Store,context.State,operation.Token));context.State.SharedVaultPath=destination;
                    }
                }
                else {BeginOperation();await Task.Run(()=>{foreach(var doc in context.State.Documents){operation.Token.ThrowIfCancellationRequested();context.Store.ResolveDocumentPath(doc);}});context.State.SharedVaultPath="";}
                try{context.Store.Save(context.State);}catch{context.State.SharedVaultPath=old;throw;}
                context.MarkWorkbookDirty();SetStatus(response==DialogResult.Yes?"Mode partagé. Enregistrez Excel pour transmettre les références.":"Mode autonome. Enregistrez Excel pour intégrer toutes les pièces.");
            }
            catch(OperationCanceledException){SetStatus("Changement de stockage annulé. Le mode précédent est conservé.");}
            catch(Exception ex){ShowError(ex);}finally{EndOperation();}
        }
    }
}
