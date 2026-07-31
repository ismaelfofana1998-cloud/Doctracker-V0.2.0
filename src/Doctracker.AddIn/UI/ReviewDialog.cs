using System.Drawing;
using System.Windows.Forms;
using Doctracker.Core.Models;

namespace Doctracker.AddIn.UI
{
    internal sealed class ReviewDialog : Form
    {
        private readonly ComboBox status;
        private readonly TextBox comment;

        public ReviewDialog(SnipRecord snip)
        {
            Text = "Revue de la preuve";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(430, 230);

            var statusLabel = new Label
            {
                Text = "Statut",
                Location = new Point(18, 18),
                AutoSize = true
            };
            status = new ComboBox
            {
                Location = new Point(18, 42),
                Width = 390,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            status.Items.AddRange(new object[]
            {
                ReviewStatus.Prepared,
                ReviewStatus.Reviewed,
                ReviewStatus.Rejected
            });
            status.SelectedItem = snip.Status;

            var commentLabel = new Label
            {
                Text = "Commentaire de revue",
                Location = new Point(18, 78),
                AutoSize = true
            };
            comment = new TextBox
            {
                Location = new Point(18, 102),
                Width = 390,
                Height = 70,
                Multiline = true,
                Text = snip.Comment
            };
            var save = new Button
            {
                Text = "Enregistrer",
                DialogResult = DialogResult.OK,
                Location = new Point(222, 186),
                Width = 90
            };
            var cancel = new Button
            {
                Text = "Annuler",
                DialogResult = DialogResult.Cancel,
                Location = new Point(318, 186),
                Width = 90
            };

            Controls.Add(statusLabel);
            Controls.Add(status);
            Controls.Add(commentLabel);
            Controls.Add(comment);
            Controls.Add(save);
            Controls.Add(cancel);
            AcceptButton = save;
            CancelButton = cancel;
        }

        public ReviewStatus SelectedStatus => (ReviewStatus)status.SelectedItem;
        public string ReviewComment => comment.Text.Trim();
    }
}
