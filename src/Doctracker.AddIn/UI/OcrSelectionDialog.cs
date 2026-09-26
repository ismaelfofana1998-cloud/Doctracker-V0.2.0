using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.UI
{
    internal sealed class OcrSelectionDialog : Form
    {
        private readonly ProjectState state;
        private readonly HashSet<string> selected=new HashSet<string>();
        private readonly ComboBox folders=new ComboBox {Dock=DockStyle.Top,DropDownStyle=ComboBoxStyle.DropDownList,DisplayMember="Caption",AccessibleName="Dossier OCR"};
        private readonly ListView files=new ListView {Dock=DockStyle.Fill,View=View.Details,CheckBoxes=true,FullRowSelect=true,MultiSelect=true,HideSelection=false,AccessibleName="Documents à reconnaître"};
        private readonly Button run=new Button {Text="Lancer l’OCR",AutoSize=true,Enabled=false,DialogResult=DialogResult.OK};
        private readonly Label summary=new Label {AutoSize=false,Height=52,Dock=DockStyle.Bottom,Padding=new Padding(0,8,0,8)};
        private bool binding;
        public List<DocumentRecord> SelectedDocuments=>state.Documents.Where(d=>selected.Contains(d.Id)).ToList();
        public OcrSelectionDialog(ProjectState state,string folder,IEnumerable<string> initiallySelected=null)
        {
            this.state=state;
            if(initiallySelected!=null)selected.UnionWith(initiallySelected);
            Text="OCR — choisir les documents";Font=new Font("Segoe UI",10);ClientSize=new Size(820,510);MinimumSize=new Size(660,400);
            StartPosition=FormStartPosition.CenterParent;Padding=new Padding(12);MinimizeBox=false;MaximizeBox=false;
            var note=new Label {Text="Cochez les documents ou choisissez un dossier. Le texte déjà préparé est réutilisé ; l’OCR traite les pages restantes.",Dock=DockStyle.Top,Height=48};
            files.Columns.Add("Document",410);files.Columns.Add("Dossier",170);files.Columns.Add("Texte",170);
            var toolbar=new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true};
            AddButton(toolbar,"Tout le dossier",()=>CheckVisible(true));
            AddButton(toolbar,"Décocher le dossier",()=>CheckVisible(false));
            AddButton(toolbar,"Cocher la sélection",()=>{foreach(ListViewItem row in files.SelectedItems)row.Checked=true;});
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Bottom,AutoSize=true,FlowDirection=FlowDirection.RightToLeft};
            var cancel=new Button {Text="Annuler",AutoSize=true,DialogResult=DialogResult.Cancel};buttons.Controls.Add(cancel);buttons.Controls.Add(run);
            Controls.Add(files);Controls.Add(toolbar);Controls.Add(folders);Controls.Add(note);Controls.Add(summary);Controls.Add(buttons);
            AcceptButton=run;CancelButton=cancel;
            files.ItemChecked+=(s,e)=>{if(binding)return;var id=((DocumentRecord)e.Item.Tag).Id;if(e.Item.Checked)selected.Add(id);else selected.Remove(id);UpdateCount();};
            folders.SelectedIndexChanged+=(s,e)=>RefreshFiles();
            var choices=new[]{new FolderChoice {Path=null,Caption="Tous les documents"},new FolderChoice {Path="",Caption="Non classés"}}
                .Concat(ProjectFolders.Paths(state).Select(path=>new FolderChoice {Path=path,Caption=path})).ToList();
            folders.DataSource=choices;folders.SelectedItem=choices.FirstOrDefault(c=>c.Path==folder)??choices[0];RefreshFiles();
        }
        private static void AddButton(FlowLayoutPanel panel,string text,Action action)
        {var button=new Button {Text=text,AutoSize=true};button.Click+=(s,e)=>action();panel.Controls.Add(button);}
        private void CheckVisible(bool check) { foreach(ListViewItem row in files.Items)row.Checked=check; }
        private void UpdateCount()
        {
            var count=SelectedDocuments.Count;run.Enabled=count>0;run.Text="Lancer l’OCR ("+count+")";
            summary.Text=count+" document(s) coché(s), y compris dans les autres dossiers. Les documents déjà prêts ne seront pas retraités.";
        }
        private void RefreshFiles()
        {
            binding=true;files.BeginUpdate();
            try
            {
                files.Items.Clear();var folder=(folders.SelectedItem as FolderChoice)?.Path;
                foreach(var doc in RecognitionScope.Select(state,folder).OrderBy(d=>d.DisplayName))
                {
                    var row=new ListViewItem(doc.DisplayName) {Tag=doc,Checked=selected.Contains(doc.Id)};
                    row.SubItems.Add(string.Join(" ; ",doc.Categories));
                    row.SubItems.Add(!string.IsNullOrEmpty(doc.IndexError)?"À réessayer":doc.IndexComplete?"Prêt":"À préparer");files.Items.Add(row);
                }
            }
            finally {files.EndUpdate();binding=false;UpdateCount();}
        }
    }
}
