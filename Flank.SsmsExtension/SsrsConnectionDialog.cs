using System;
using System.Drawing;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class SsrsConnectionOptions
    {
        public string ReportServerUrl { get; set; }
        public string ReportPortalUrl { get; set; }
    }

    internal sealed class SsrsConnectionDialog : Form
    {
        private readonly TextBox reportServerTextBox;
        private readonly TextBox reportPortalTextBox;

        public SsrsConnectionOptions Options { get; private set; }

        public SsrsConnectionDialog()
        {
            Text = "Connect to SSRS";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(540, 150);

            reportServerTextBox = AddField(
                "Report Server URL:",
                "http://localhost/ReportServer",
                20);

            reportPortalTextBox = AddField(
                "Report Portal URL:",
                "http://localhost/Reports",
                60);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(350, 105),
                Width = 80
            };

            var continueButton = new Button
            {
                Text = "Continue",
                Location = new Point(440, 105),
                Width = 80
            };

            continueButton.Click += ContinueButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(continueButton);

            AcceptButton = continueButton;
            CancelButton = cancelButton;
        }

        private TextBox AddField(
            string label,
            string defaultValue,
            int y)
        {
            var labelControl = new Label
            {
                Text = label,
                Location = new Point(20, y + 4),
                Width = 130
            };

            var textBox = new TextBox
            {
                Text = defaultValue,
                Location = new Point(155, y),
                Width = 365
            };

            Controls.Add(labelControl);
            Controls.Add(textBox);

            return textBox;
        }

        private void ContinueButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(reportServerTextBox.Text) ||
                string.IsNullOrWhiteSpace(reportPortalTextBox.Text))
            {
                MessageBox.Show(
                    "Both URLs are required.",
                    "Flank",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            Options = new SsrsConnectionOptions
            {
                ReportServerUrl = reportServerTextBox.Text.Trim(),
                ReportPortalUrl = reportPortalTextBox.Text.Trim()
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}