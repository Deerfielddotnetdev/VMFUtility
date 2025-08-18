using System;
using System.Drawing;
using System.Windows.Forms;

namespace MailFlowSuperUtility
{
    public class UnlockForm : Form
    {
        // >>> CHANGE THIS PER BUILD (e.g., invoice number)
        private const string UNLOCK_KEY = "CHANGE_ME";

        private TextBox txtUnlockKey;
        private Button btnActivate;
        private Button btnCancel;

        public UnlockForm()
        {
            Text = "Enter Unlock Code";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(420, 140);

            var lblHeader = new Label
            {
                AutoSize = true,
                Location = new Point(12, 12),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Text = "Enter Unlock Code (Invoice #)"
            };
            Controls.Add(lblHeader);

            var lblKey = new Label
            {
                AutoSize = true,
                Location = new Point(12, 50),
                Text = "Unlock Code:"
            };
            Controls.Add(lblKey);

            txtUnlockKey = new TextBox
            {
                Location = new Point(100, 46),
                Width = 300
            };
            Controls.Add(txtUnlockKey);

            btnActivate = new Button
            {
                Text = "Activate",
                Location = new Point(244, 90),
                DialogResult = DialogResult.None
            };
            btnActivate.Click += BtnActivate_Click;
            Controls.Add(btnActivate);

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(325, 90),
                DialogResult = DialogResult.Cancel
            };
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnCancel);

            AcceptButton = btnActivate;
            CancelButton = btnCancel;
        }

        private void BtnActivate_Click(object sender, EventArgs e)
        {
            var entered = (txtUnlockKey.Text ?? string.Empty).Trim();
            if (string.Equals(entered, UNLOCK_KEY, StringComparison.Ordinal))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(this, "Invalid unlock code. Please try again.", "Invalid",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUnlockKey.SelectAll();
                txtUnlockKey.Focus();
            }
        }
    }
}
