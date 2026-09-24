using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.UI
{
    internal sealed class FolderChoice
    {
        public string Path {get;set;}
        public string Caption {get;set;}
        public override string ToString()=>Caption;
    }
    internal sealed class FolderOrganizer : Form
    {
        private const string DragFormat="Doctracker.DocumentIds";
        private readonly ProjectState state;
        private readonly ProjectFolders folders;
        private readonly TreeView tree=new TreeView {Dock=DockStyle.Fill,HideSelection=false,AllowDrop=true,BorderStyle=BorderStyle.None};
        private readonly ListView files=new ListView {Dock=DockStyle.Fill,View=View.Details,FullRowSelect=true,MultiSelect=true,HideSelection=false,BorderStyle=BorderStyle.None};
        private readonly Label info=new Label {Dock=DockStyle.Bottom,AutoSize=true,Padding=new Padding(10),Text="Sélectionnez les documents (Ctrl / Maj), puis glissez-les sur un dossier à gauche."};
        public string SelectedDocumentId {get;private set;}
        public FolderOrganizer(ProjectStore store,ProjectState state)
        {
            this.state=state;folders=new ProjectFolders(store);
            Text="Dossiers de travail";Font=new Font("Segoe UI",10);ClientSize=new Size(880,530);MinimumSize=new Size(640,380);StartPosition=FormStartPosition.CenterParent;
            BackColor=Color.White;
            var toolbar=new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true,Padding=new Padding(6)};
            var create=SnipTheme.Button("Nouveau dossier","ImportDocuments");create.Click+=(s,e)=>Run(()=>{
                var name=Ask("Nouveau dossier","");if(name==null)return;var parent=(tree.SelectedNode?.Tag as FolderChoice)?.Path??"";
                folders.Create(state,parent,name);RefreshTree(string.IsNullOrEmpty(parent)?name.Trim():parent+" / "+name.Trim());});
            var rename=SnipTheme.Button("Renommer");rename.Click+=(s,e)=>Run(()=>{
                var path=SelectedPath;if(string.IsNullOrEmpty(path))return;var name=Ask("Renommer le dossier",path.Split(new[]{" / "},StringSplitOptions.None).Last());if(name==null)return;
                folders.Rename(state,path,name);RefreshTree(ProjectFolders.Parent(path)==""?name.Trim():ProjectFolders.Parent(path)+" / "+name.Trim());});
            var remove=SnipTheme.Button("Supprimer le dossier");remove.Click+=(s,e)=>Run(()=>{
                var path=SelectedPath;if(string.IsNullOrEmpty(path))return;
                if(MessageBox.Show(this,"Supprimer ce dossier et ses sous-dossiers ? Les documents seront conservés et remontés au dossier parent.","Dossier",MessageBoxButtons.YesNo)!=DialogResult.Yes)return;
                folders.Remove(state,path);RefreshTree(ProjectFolders.Parent(path));});
            var move=SnipTheme.Button("Déplacer vers…");move.Click+=(s,e)=>{
                var ids=SelectedIds();if(ids.Length==0){info.Text="Sélectionnez au moins un document dans la liste.";return;}
                var menu=new ContextMenuStrip();menu.Items.Add("Non classés",null,(a,b)=>Move(ids,""));
                foreach(var path in ProjectFolders.Paths(state)){var target=path;menu.Items.Add(target,null,(a,b)=>Move(ids,target));}
                menu.Closed+=(a,b)=>menu.Dispose();menu.Show(move,new Point(0,move.Height));};
            toolbar.Controls.Add(create);toolbar.Controls.Add(rename);toolbar.Controls.Add(remove);toolbar.Controls.Add(move);
            var split=new SplitContainer {Size=new Size(850,400),Dock=DockStyle.Fill,FixedPanel=FixedPanel.Panel1,SplitterDistance=230};
            split.Panel1.Controls.Add(tree);split.Panel2.Controls.Add(files);files.Columns.Add("Document",400);files.Columns.Add("Dossier",240);
            Controls.Add(split);Controls.Add(info);Controls.Add(toolbar);
            tree.AfterSelect+=(s,e)=>RefreshFiles();
            files.ItemDrag+=(s,e)=>{var ids=SelectedIds();if(ids.Length>0)files.DoDragDrop(new DataObject(DragFormat,ids),DragDropEffects.Move);};
            tree.DragOver+=(s,e)=>{var node=tree.GetNodeAt(tree.PointToClient(new Point(e.X,e.Y)));e.Effect=e.Data.GetDataPresent(DragFormat)&&node?.Tag is FolderChoice choice && choice.Path!=null?DragDropEffects.Move:DragDropEffects.None;};
            tree.DragDrop+=(s,e)=>{var node=tree.GetNodeAt(tree.PointToClient(new Point(e.X,e.Y)));var destination=(node?.Tag as FolderChoice)?.Path;
                if(destination!=null && e.Data.GetData(DragFormat) is string[] ids)Move(ids,destination);};
            files.DoubleClick+=(s,e)=>{if(files.SelectedItems.Count==1){SelectedDocumentId=(string)files.SelectedItems[0].Tag;DialogResult=DialogResult.OK;Close();}};
            RefreshTree(null);
        }
        private string SelectedPath=>(tree.SelectedNode?.Tag as FolderChoice)?.Path;
        private string[] SelectedIds()=>files.SelectedItems.Cast<ListViewItem>().Select(i=>(string)i.Tag).ToArray();
        private void Move(string[] ids,string path)=>Run(()=>{folders.Move(state,ids,path);RefreshTree(path);info.Text=ids.Length+" document(s) déplacé(s). Les liens et les snips sont conservés.";});
        private void RefreshTree(string selected)
        {
            tree.BeginUpdate();tree.Nodes.Clear();
            var all=new TreeNode("Tous les documents") {Tag=new FolderChoice {Path=null}};
            tree.Nodes.Add(all);var unfiled=new TreeNode("Non classés") {Tag=new FolderChoice {Path=""}};tree.Nodes.Add(unfiled);
            var nodes=new System.Collections.Generic.Dictionary<string,TreeNode>(StringComparer.OrdinalIgnoreCase);TreeNode active=selected==null?all:unfiled;
            foreach(var path in ProjectFolders.Paths(state))
            {
                var node=new TreeNode(path.Split(new[]{" / "},StringSplitOptions.None).Last()) {Tag=new FolderChoice {Path=path}};nodes[path]=node;
                var parent=ProjectFolders.Parent(path);if(parent.Length>0 && nodes.TryGetValue(parent,out var parentNode))parentNode.Nodes.Add(node);else tree.Nodes.Add(node);
                if(string.Equals(path,selected,StringComparison.OrdinalIgnoreCase))active=node;
            }
            tree.ExpandAll();tree.SelectedNode=active;tree.EndUpdate();RefreshFiles();
        }
        private void RefreshFiles()
        {
            var path=SelectedPath;files.BeginUpdate();files.Items.Clear();
            foreach(var doc in state.Documents.Where(d=>path==null || (path==""?d.Categories.Count==0:d.Categories.Any(c=>ProjectFolders.Within(c,path)))).OrderBy(d=>d.DisplayName))
            {var row=new ListViewItem(doc.DisplayName){Tag=doc.Id};row.SubItems.Add(string.Join(" ; ",doc.Categories));files.Items.Add(row);}
            files.EndUpdate();
        }
        private void Run(Action action){try{action();}catch(Exception ex){MessageBox.Show(this,ex.GetBaseException().Message,"Dossiers",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
        private string Ask(string title,string value)
        {
            using(var dialog=new Form {Text=title,ClientSize=new Size(390,105),StartPosition=FormStartPosition.CenterParent,Font=Font,Padding=new Padding(10)})
            {var input=new TextBox {Text=value,Dock=DockStyle.Top,MaxLength=80};var ok=new Button {Text="Enregistrer",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK};dialog.Controls.Add(input);dialog.Controls.Add(ok);dialog.AcceptButton=ok;
                return dialog.ShowDialog(this)==DialogResult.OK?input.Text:null;}
        }
    }
}
