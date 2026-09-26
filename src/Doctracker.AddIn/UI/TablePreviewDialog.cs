using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.UI
{
    internal sealed class TablePreviewDialog : Form
    {
        private readonly DocumentCanvas canvas=new DocumentCanvas();
        private readonly DataGridView grid;
        private readonly IReadOnlyList<WordRecord> words;
        private readonly Dictionary<string,string> corrections=new Dictionary<string,string>();
        private readonly Label summary;
        public Exception Failure {get;private set;}
        public TableGrid GridDefinition {get;private set;}
        public IReadOnlyList<ExtractedCell> Cells {get;private set;}
        public int RowCount=>GridDefinition.RowCount;
        public int ColumnCount=>GridDefinition.ColumnCount;
        public TablePreviewDialog(string path,int page,RectangleF zone,IReadOnlyList<WordRecord> words)
        {
            this.words=words;GridDefinition=TableGrid.FromWords(words);
            Text="Tableau · tracer les lignes et les colonnes";
            ClientSize=new Size(1100,760);MinimumSize=new Size(820,600);StartPosition=FormStartPosition.CenterParent;
            Font=new Font("Segoe UI",9);AutoScaleMode=AutoScaleMode.Dpi;
            var split=new SplitContainer {Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,Size=new Size(1100,650),
                SplitterDistance=460,Panel1MinSize=180,Panel2MinSize=110};
            grid=new DataGridView {Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,
                AllowUserToOrderColumns=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,RowHeadersVisible=true,
                BackgroundColor=Color.White,SelectionMode=DataGridViewSelectionMode.CellSelect,MultiSelect=false};
            split.Panel1.Controls.Add(canvas);split.Panel2.Controls.Add(grid);
            var tools=new FlowLayoutPanel {Dock=DockStyle.Top,AutoSize=true,Padding=new Padding(8,4,8,4),WrapContents=true};
            var move=new RadioButton {Text="Points",Checked=true,AutoSize=true,Appearance=Appearance.Button};
            var column=new RadioButton {Text="+ Colonne",AutoSize=true,Appearance=Appearance.Button};
            var row=new RadioButton {Text="+ Ligne",AutoSize=true,Appearance=Appearance.Button};
            move.CheckedChanged+=(s,e)=>{if(move.Checked)canvas.TableAddColumn=null;};
            column.CheckedChanged+=(s,e)=>{if(column.Checked)canvas.TableAddColumn=true;};
            row.CheckedChanged+=(s,e)=>{if(row.Checked)canvas.TableAddColumn=false;};
            tools.Controls.AddRange(new Control[]{move,column,row});
            summary=new Label {AutoSize=true,Padding=new Padding(12,7,0,0)};tools.Controls.Add(summary);
            var guide=new Label {Dock=DockStyle.Top,Height=46,Padding=new Padding(10,4,10,4),Text=
                "Cliquez en haut du cadre pour ajouter une colonne, à gauche pour une ligne. Déplacez les points ; clic droit pour les supprimer.\nCorrigez si besoin les valeurs dans l’aperçu en dessous. Ctrl + molette pour zoomer."};
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Bottom,AutoSize=true,Padding=new Padding(8),FlowDirection=FlowDirection.RightToLeft};
            var insert=new Button {Text="Insérer dans Excel",AutoSize=true};
            insert.Click+=(s,e)=>{grid.EndEdit();DialogResult=DialogResult.OK;Close();};
            var cancel=new Button {Text="Annuler",AutoSize=true,DialogResult=DialogResult.Cancel};
            buttons.Controls.AddRange(new Control[]{insert,cancel});
            Controls.Add(split);Controls.Add(guide);Controls.Add(tools);Controls.Add(buttons);
            CancelButton=cancel; // Enter remains available for editing preview cells.
            canvas.TableGridChanged+=RebuildPreview;
            RebuildPreview();
            canvas.DisplayFailed+=failure=>{Failure=failure;DialogResult=DialogResult.Cancel;Close();};
            Shown+=(s,e)=>{
                try {canvas.LoadDocument(path);if(Failure==null)canvas.EditTableGrid(page,zone,GridDefinition);}
                catch(Exception failure){Failure=failure;DialogResult=DialogResult.Cancel;Close();}
            };
        }
        private static string Key(ExtractedCell cell)=>string.Join("|",new[]{cell.X,cell.Y,cell.Width,cell.Height}.Select(v=>v.ToString("R",CultureInfo.InvariantCulture)));
        private void RebuildPreview()
        {
            grid.EndEdit();
            if(Cells!=null)foreach(var cell in Cells)
            {
                var value=ValueAt(cell.Row,cell.Column);
                if(value!=cell.Text)corrections[Key(cell)]=value;else corrections.Remove(Key(cell));
            }
            Cells=GridDefinition.Extract(words);
            grid.Rows.Clear();grid.Columns.Clear();
            for(var i=0;i<ColumnCount;i++)grid.Columns.Add(new DataGridViewTextBoxColumn {Name="c"+i,HeaderText="Colonne "+(i+1),SortMode=DataGridViewColumnSortMode.NotSortable});
            grid.Rows.Add(RowCount);
            foreach(var cell in Cells)grid.Rows[cell.Row].Cells[cell.Column].Value=corrections.TryGetValue(Key(cell),out var value)?value:cell.Text;
            for(var i=0;i<RowCount;i++)grid.Rows[i].HeaderCell.Value=(i+1).ToString();
            summary.Text=RowCount+" ligne(s) · "+ColumnCount+" colonne(s)";
        }
        public string ValueAt(int row,int column)=>Convert.ToString(grid.Rows[row].Cells[column].Value);
    }
}
