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
        public readonly ComboBox Documents = new EvidenceComboBox { DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "DisplayName", FlatStyle = FlatStyle.Standard, Dock = DockStyle.Fill, IntegralHeight = false, DropDownHeight = 320, AccessibleName = "Document actif" };
        public readonly TextBox Query = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Rechercher dans les documents" };
        public readonly ListBox Results = new ListBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, IntegralHeight = false, DisplayMember = "Caption", HorizontalScrollbar = true, AccessibleName = "Résultats de recherche" };
        public readonly ComboBox Proofs = new EvidenceComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = "Caption", FlatStyle = FlatStyle.Flat, AccessibleName = "Preuves de la cellule" };
        public readonly Label Brand = Label("Doctracker");
        private readonly ProgressBar progress = new ProgressBar {Height=4,Dock=DockStyle.Fill,Style=ProgressBarStyle.Marquee,MarqueeAnimationSpeed=30,Visible=false};
        public readonly Label Status = Label("Prêt");
        public readonly Button Search = SnipTheme.Button("Rechercher", "SearchDocuments");
        public readonly Button DeleteSnip = SnipTheme.Button("", "Delete", Color.Red);
        public readonly Button Cancel = SnipTheme.Button("Annuler");
        public readonly ComboBox Categories = new EvidenceComboBox { DropDownStyle=ComboBoxStyle.DropDownList, Dock=DockStyle.Fill, FlatStyle=FlatStyle.Flat, AccessibleName="Dossier de documents" };
        public readonly DocumentCanvas Canvas = new DocumentCanvas();
        private readonly TableLayoutPanel resultsPanel;
        private readonly TableLayoutPanel proofPanel;
        private readonly Label resultCount = Label("");
        private readonly ToolTip tips = new ToolTip();
        private int resultTotal;

        public WorkspaceView()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96,96);
            Font = new Font("Segoe UI",9F);
            BackColor = Color.White; ForeColor = SnipTheme.Ink; Dock = DockStyle.Fill;
            Size = new Size(760,700);
            var root = Rows(); root.AutoSize=false; root.Dock = DockStyle.Fill;
            var header = Row(0,100); header.ColumnStyles[0] = new ColumnStyle(SizeType.AutoSize); header.Padding = new Padding(8,6,8,4);
            header.BackColor=SnipTheme.Surface;
            Brand.ForeColor=SnipTheme.Ink;Brand.Padding=new Padding(0,0,12,0);Brand.Anchor=AnchorStyles.Left;Brand.Dock=DockStyle.None;
            header.Controls.Add(Brand,0,0);
            var searchRow=Row(100,0);searchRow.ColumnStyles[1]=new ColumnStyle(SizeType.AutoSize);
            Query.Anchor=AnchorStyles.Left|AnchorStyles.Right;Query.Margin=new Padding(0,0,2,0);
            Search.Text="";Search.AccessibleName="Rechercher";Search.Padding=new Padding(3,2,3,2);Search.Margin=Padding.Empty;
            searchRow.Controls.Add(Query,0,0);searchRow.Controls.Add(Search,1,0);header.Controls.Add(searchRow,1,0);Add(root,header);
            tips.SetToolTip(Query,"Rechercher un texte, une référence, une date ou un montant dans le dossier actif (Entrée)");
            tips.SetToolTip(Search,"Rechercher dans le dossier actif");
            tips.SetToolTip(Documents,"Document actif · ouvrez la liste pour afficher les noms complets");
            Documents.DropDown += (s,e) => {
                var longest=Documents.Items.Cast<object>().Select(item=>Documents.GetItemText(item)).Aggregate("",(a,b)=>b.Length>a.Length?b:a);
                Documents.DropDownWidth=Math.Max(Documents.Width,Math.Min(900,TextRenderer.MeasureText(longest,Documents.Font).Width+40));
            };
            var filters=Row(35,65);filters.AutoSize=false;filters.Padding=new Padding(8,2,8,4);
            Categories.Margin=new Padding(0,0,8,0);Documents.Margin=Padding.Empty;
            Categories.Anchor=Documents.Anchor=AnchorStyles.Left|AnchorStyles.Right;
            filters.Controls.Add(Categories,0,0);filters.Controls.Add(Documents,1,0);
            Add(root,filters);
            Categories.Items.Add("Tous les documents");Categories.SelectedIndex=0;
            tips.SetToolTip(Categories,"Dossier actif : il limite la recherche et l'export");
            DeleteSnip.AccessibleName="Supprimer le snip sélectionné";
            resultsPanel = Rows();resultsPanel.Padding=new Padding(12,0,12,8);resultsPanel.Visible=false;
            var resultsHeader=Row(100,0);resultsHeader.ColumnStyles[1]=new ColumnStyle(SizeType.AutoSize);
            var close=SnipTheme.Button("Masquer");close.Click+=(s,e)=>resultsPanel.Visible=false;
            resultsHeader.Controls.Add(resultCount,0,0);resultsHeader.Controls.Add(close,1,0);Add(resultsPanel,resultsHeader);
            Results.Height=76;Add(resultsPanel,Results,SizeType.Absolute,76);Add(root,resultsPanel);
            Results.FontChanged+=(s,e)=>SizeResults();
            proofPanel=Row(0,100);proofPanel.ColumnStyles[0]=new ColumnStyle(SizeType.AutoSize);proofPanel.Padding=new Padding(10,5,10,5);proofPanel.Visible=false;
            proofPanel.Controls.Add(Label("Preuve"),0,0);proofPanel.Controls.Add(Proofs,1,0);proofPanel.ColumnCount=3;proofPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));proofPanel.Controls.Add(DeleteSnip,2,0);Add(root,proofPanel);
            Add(root,Canvas,SizeType.Percent,100);
            Add(root,progress);
            var footer=Row(100,0);footer.ColumnStyles[1]=new ColumnStyle(SizeType.AutoSize);footer.BackColor=SnipTheme.Surface;
            Status.AutoSize=false;Status.AutoEllipsis=true;Status.Padding=new Padding(8,2,8,2);Status.Height=Font.Height+8;
            Status.FontChanged+=(s,e)=>Status.Height=Status.Font.Height+8;
            Status.TextChanged+=(s,e)=>tips.SetToolTip(Status,Status.Text);
            Cancel.Visible=false;Cancel.Margin=Padding.Empty;Cancel.Padding=new Padding(4,0,4,0);
            footer.Controls.Add(Status,0,0);footer.Controls.Add(Cancel,1,0);Add(root,footer);
            Action sizeHeader = () => {
                var height=Math.Max(Query.PreferredHeight,Search.GetPreferredSize(Size.Empty).Height)+header.Padding.Vertical;
                root.RowStyles[0].SizeType=SizeType.Absolute;
                if(root.RowStyles[0].Height!=height)root.RowStyles[0].Height=height;
                var selectors=Math.Max(Documents.GetPreferredSize(Size.Empty).Height,Categories.GetPreferredSize(Size.Empty).Height)+filters.Padding.Vertical;
                root.RowStyles[1].SizeType=SizeType.Absolute;
                if(root.RowStyles[1].Height!=selectors)root.RowStyles[1].Height=selectors;
                var footerHeight=Math.Max(Status.Font.Height+8,Cancel.GetPreferredSize(Size.Empty).Height);
                var lastRow=root.RowStyles[root.RowCount-1];lastRow.SizeType=SizeType.Absolute;
                if(lastRow.Height!=footerHeight)lastRow.Height=footerHeight;
            };
            Documents.FontChanged+=(s,e)=>sizeHeader();Categories.FontChanged+=(s,e)=>sizeHeader();Status.FontChanged+=(s,e)=>sizeHeader();header.Layout+=(s,e)=>sizeHeader();sizeHeader();
            Controls.Add(root);
        }
        public void SetMode(SnipType? type)
        {
            Canvas.CommentMode=false;Canvas.ActiveType=type;

        }
        public void SetCommentMode(bool active)
        {
            SetMode(null);Canvas.CommentMode=active;

        }
        public void SetDocumentSummary(int total, int indexed) => tips.SetToolTip(Documents,
            "Document actif · " + total + " pièce(s), " + indexed + " indexée(s). Texte préparé à la première recherche.");
        public void ShowResults(int count) { resultTotal=count;resultCount.Text=count==0 ? "Aucun résultat · Vérifiez le dossier et le texte recherché." : count+" résultat(s)";SizeResults();resultsPanel.Visible=true; }
        private void SizeResults() { Results.Visible=resultTotal>0;resultsPanel.RowStyles[1].Height=resultTotal==0 ? 0 : Math.Min(3,resultTotal)*Results.ItemHeight+8; resultCount.MaximumSize=new Size(Math.Max(100,Width-120),0); }
        public void HideResults() { resultsPanel.Visible=false; }
        public void ShowProofs(bool visible) { proofPanel.Visible=visible; }
        public void SetBusy(bool busy) { progress.Visible=busy; Documents.Enabled=Query.Enabled=Search.Enabled=Proofs.Enabled=Categories.Enabled=DeleteSnip.Enabled=!busy;Cancel.Visible=busy; }
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
        protected override void Dispose(bool disposing) { if(disposing){tips.Dispose();}base.Dispose(disposing); }
    }
    internal sealed class EvidenceComboBox : ComboBox
    {
        public EvidenceComboBox()
        {
            DrawMode=DrawMode.OwnerDrawFixed;
            BackColor=Color.White; ForeColor=SnipTheme.Ink;
            SizeItems();
        }
        public override Size GetPreferredSize(Size proposedSize)
        {
            var size=base.GetPreferredSize(proposedSize);
            // Owner-drawn native items are taller than ComboBox.PreferredHeight (font-only).
            size.Height=Math.Max(size.Height,ItemHeight+8);
            return size;
        }
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            SizeItems();
        }
        private void SizeItems()
        {
            // Native edit/item height can remain cached after inherited font changes.
            // Measure from the current font for both the closed field and the list.
            ItemHeight=Math.Max(22,Font.Height+8);
            Height=PreferredHeight;
            Parent?.PerformLayout(this,"PreferredSize");
        }
        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            var selected=(e.State & DrawItemState.Selected)!=0;
            var edit=(e.State & DrawItemState.ComboBoxEdit)!=0;
            using(var brush=new SolidBrush(selected && !edit ? SnipTheme.Surface : Color.White)) e.Graphics.FillRectangle(brush,e.Bounds);
            if(e.Index<0 || e.Index>=Items.Count) return;
            var bounds=e.Bounds; bounds.Inflate(-5,0);
            TextRenderer.DrawText(e.Graphics,GetItemText(Items[e.Index]),Font,bounds,
                Enabled ? SnipTheme.Ink : SnipTheme.Muted,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.EndEllipsis|TextFormatFlags.NoPrefix);
        }
    }

}
