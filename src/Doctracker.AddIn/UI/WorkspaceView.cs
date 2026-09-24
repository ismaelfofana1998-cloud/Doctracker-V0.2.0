using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Doctracker.Core.Models;

namespace Doctracker.AddIn.UI
{
    // Excel-independent view: layout can be rendered and checked on Windows without COM.
    internal sealed class WorkspaceView : UserControl
    {
        public readonly ComboBox Documents = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "OriginalName", FlatStyle = FlatStyle.Flat, Dock = DockStyle.Fill, IntegralHeight = false, DropDownHeight = 320, AccessibleName = "Document actif" };
        public readonly TextBox Query = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Rechercher dans les documents" };
        public readonly ListBox Results = new ListBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, IntegralHeight = false, DisplayMember = "Caption", HorizontalScrollbar = true, AccessibleName = "Résultats de recherche" };
        public readonly ComboBox Proofs = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "Caption", FlatStyle = FlatStyle.Flat, AccessibleName = "Preuves de la cellule" };
        public readonly Label IndexState = Label("Aucune pièce");
        public readonly Label ModeState = Label("Choisissez un snip, puis dessinez une zone.");
        public readonly Label Status = Label("Prêt · Cliquez sur une cellule liée pour afficher sa preuve.");
        public readonly Button Import = SnipTheme.Button("Importer", "ImportDocuments");
        public readonly Button Search = SnipTheme.Button("Rechercher", "SearchDocuments");
        public readonly Button Cancel = SnipTheme.Button("Annuler");
        public readonly ToolStripMenuItem Reindex = new ToolStripMenuItem("Réindexer les documents");
        public readonly ToolStripMenuItem Remove = new ToolStripMenuItem("Retirer le document sélectionné");
        public readonly DocumentCanvas Canvas = new DocumentCanvas();
        private readonly TableLayoutPanel resultsPanel;
        private readonly TableLayoutPanel proofPanel;
        private readonly Label resultCount = Label("");
        private readonly FlowLayoutPanel modes;
        private readonly Dictionary<SnipType, Button> modeButtons = new Dictionary<SnipType, Button>();
        private readonly ContextMenuStrip menu = new ContextMenuStrip();
        private readonly ToolTip tips = new ToolTip();
        public event Action<SnipType?> ModeChanged;
        private SnipType? activeMode;

        public WorkspaceView()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96,96);
            Font = new Font("Segoe UI",9F);
            BackColor = Color.White; ForeColor = SnipTheme.Ink; Dock = DockStyle.Fill;
            Size = new Size(760,700);
            var root = Rows(); root.AutoSize=false; root.Dock = DockStyle.Fill;
            var header = Row(38,62); header.Padding = new Padding(10,10,10,4);
            var brand = Label("Doctracker"); brand.Font = new Font("Segoe UI Semibold",14,FontStyle.Bold); brand.ForeColor=SnipTheme.Ink;
            brand.Anchor=AnchorStyles.Left; header.Controls.Add(brand,0,0);
            var picker = Rows(); picker.Dock=DockStyle.Fill;
            var caption = Label("DOCUMENTS"); caption.Font=new Font("Segoe UI",8F,FontStyle.Bold);
            Add(picker,caption); Add(picker,Documents); header.Controls.Add(picker,1,0); Add(root,header);
            Documents.DropDown += (s,e) => { Documents.DropDownWidth = Math.Max(Documents.Width, Math.Min(900, Documents.Items.Cast<object>().Select(item=>TextRenderer.MeasureText(Convert.ToString(item.GetType().GetProperty("OriginalName")?.GetValue(item,null) ?? item),Documents.Font).Width+40).DefaultIfEmpty(300).Max())); };
            var actions = new FlowLayoutPanel { AutoSize=true, Dock=DockStyle.Fill, Padding=new Padding(8,0,8,4), Margin=Padding.Empty };
            var more = SnipTheme.Button("", "More"); more.AccessibleName="Options des documents";
            menu.Items.Add(Reindex); menu.Items.Add(Remove); more.Click+=(s,e)=>menu.Show(more,new Point(0,more.Height));
            tips.SetToolTip(more,"Réindexer ou retirer un document");
            IndexState.Margin=new Padding(8,10,0,0);
            actions.Controls.Add(Import);actions.Controls.Add(more);actions.Controls.Add(IndexState);Add(root,actions);
            var searchRow = Row(100,0);searchRow.ColumnStyles[1]=new ColumnStyle(SizeType.AutoSize);searchRow.Padding=new Padding(10,0,10,6);
            Query.Anchor=AnchorStyles.Left|AnchorStyles.Right;searchRow.Controls.Add(Query,0,0);searchRow.Controls.Add(Search,1,0);Add(root,searchRow);
            resultsPanel = Rows();resultsPanel.Padding=new Padding(12,0,12,8);resultsPanel.Visible=false;
            var resultsHeader=Row(100,0);resultsHeader.ColumnStyles[1]=new ColumnStyle(SizeType.AutoSize);
            var close=SnipTheme.Button("Masquer");close.Click+=(s,e)=>resultsPanel.Visible=false;
            resultsHeader.Controls.Add(resultCount,0,0);resultsHeader.Controls.Add(close,1,0);Add(resultsPanel,resultsHeader);
            Results.Height=76;Add(resultsPanel,Results,SizeType.Absolute,76);Add(root,resultsPanel);
            modes = new FlowLayoutPanel {AutoSize=true,Dock=DockStyle.Fill,WrapContents=true,Padding=new Padding(8,4,8,4),Margin=Padding.Empty,BackColor=SnipTheme.Surface};
            foreach(SnipType type in new[]{SnipType.Text,SnipType.Number,SnipType.Date,SnipType.Sum,SnipType.Table,SnipType.Validation,SnipType.Exception})
            {
                var button=SnipTheme.Button(SnipTheme.LabelFor(type),type.ToString(),SnipTheme.ColorFor(type));
                button.Click+=(s,e)=>ModeChanged?.Invoke(activeMode==type ? (SnipType?)null:type);
                tips.SetToolTip(button,SnipTheme.LabelFor(type)+" : dessiner une zone. Recliquez pour désactiver.");
                modeButtons.Add(type,button);modes.Controls.Add(button);
            }
            Add(root,modes);
            ModeState.Padding=new Padding(12,5,12,7);ModeState.BackColor=SnipTheme.Surface;Add(root,ModeState);
            proofPanel=Row(0,100);proofPanel.ColumnStyles[0]=new ColumnStyle(SizeType.AutoSize);proofPanel.Padding=new Padding(10,5,10,5);proofPanel.Visible=false;
            proofPanel.Controls.Add(Label("Preuve"),0,0);proofPanel.Controls.Add(Proofs,1,0);Add(root,proofPanel);
            Add(root,Canvas,SizeType.Percent,100);
            Cancel.Visible=false;Add(root,Cancel);
            Status.Padding=new Padding(12,8,12,8);Status.BackColor=SnipTheme.Surface;Add(root,Status);
            Controls.Add(root);
            Resize+=(s,e)=> { var w=Math.Max(100,ClientSize.Width); ModeState.MaximumSize=new Size(w,0);Status.MaximumSize=new Size(w,0);IndexState.MaximumSize=new Size(Math.Max(100,w-155),0); modes.MaximumSize=new Size(w,0); };
        }
        public void SetMode(SnipType? type)
        {
            activeMode=type;Canvas.ActiveType=type;
            foreach(var pair in modeButtons){ pair.Value.BackColor=type==pair.Key?SnipTheme.Tint(pair.Key):Color.White; pair.Value.FlatAppearance.BorderSize=type==pair.Key?1:0;pair.Value.FlatAppearance.BorderColor=SnipTheme.ColorFor(pair.Key); }
            ModeState.Text=type.HasValue ? SnipTheme.LabelFor(type.Value)+" · Dessinez une zone. Le mode reste actif." : "Choisissez un snip, puis dessinez une zone.";
            ModeState.ForeColor=type.HasValue?SnipTheme.ColorFor(type.Value):SnipTheme.Muted;
        }
        public void ShowResults(int count) { resultCount.Text=count+" résultat(s)";resultsPanel.Visible=true; }
        public void HideResults() { resultsPanel.Visible=false; }
        public void ShowProofs(bool visible) { proofPanel.Visible=visible; }
        public void SetBusy(bool busy) { modes.Enabled=Import.Enabled=Search.Enabled=Proofs.Enabled=!busy;Cancel.Visible=busy; }
        private static Label Label(string text) => new Label {Text=text,AutoSize=true,Dock=DockStyle.Fill,ForeColor=SnipTheme.Muted,Margin=Padding.Empty,Padding=new Padding(0,3,0,3)};
        private static TableLayoutPanel Rows()
        {
            var table=new TableLayoutPanel {AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=1,RowCount=0,Dock=DockStyle.Fill,Margin=Padding.Empty};
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));return table;
        }
        private static TableLayoutPanel Row(float left,float right)
        {
            var row=Rows();row.ColumnCount=2;row.RowCount=1;row.ColumnStyles.Clear();row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,left));row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,right));return row;
        }
        private static void Add(TableLayoutPanel table, Control control, SizeType sizing=SizeType.AutoSize,float height=0)
        { var row=table.RowCount++;table.RowStyles.Add(new RowStyle(sizing,height));table.Controls.Add(control,0,row); }
        protected override void Dispose(bool disposing) { if(disposing){menu.Dispose();tips.Dispose();}base.Dispose(disposing); }
    }
}
