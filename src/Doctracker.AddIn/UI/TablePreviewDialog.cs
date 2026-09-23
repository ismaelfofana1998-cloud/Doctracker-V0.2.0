using System.Collections.Generic;
using System.Windows.Forms;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.UI
{
    internal sealed class TablePreviewDialog : Form
    {
        private readonly DataGridView grid;
        public TablePreviewDialog(IReadOnlyList<ExtractedCell> cells, int rows, int columns)
        {
            Text = "Vérifier le tableau avant insertion";
            Width = 800; Height = 480; StartPosition = FormStartPosition.CenterParent;
            grid = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = true };
            for (var index = 0; index < columns; index++) grid.Columns.Add("c" + index, "Colonne " + (index + 1));
            grid.Rows.Add(rows);
            foreach (var cell in cells) grid.Rows[cell.Row].Cells[cell.Column].Value = cell.Text;
            var label = new Label { Text = "Vérifiez les colonnes et corrigez les valeurs. Chaque cellule insérée gardera un lien vers la pièce.", Dock = DockStyle.Top, Height = 38 };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 40, FlowDirection = FlowDirection.RightToLeft };
            var insert = new Button { Text = "Insérer", DialogResult = DialogResult.OK };
            insert.Click += (sender, args) => grid.EndEdit();
            var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel };
            buttons.Controls.Add(insert); buttons.Controls.Add(cancel);
            Controls.Add(grid); Controls.Add(buttons); Controls.Add(label);
            AcceptButton = insert; CancelButton = cancel;
        }
        public string ValueAt(int row, int column) => System.Convert.ToString(grid.Rows[row].Cells[column].Value);
    }
}
