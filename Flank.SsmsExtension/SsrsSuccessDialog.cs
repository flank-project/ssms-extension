using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class SsrsSuccessDialog : Form
    {
        private readonly string reportUrl;

        public SsrsSuccessDialog(
            string reportName,
            string reportUrl)
        {
            this.reportUrl = reportUrl;

            Text = "SSRS Report Created";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 175);

            var successLabel = new Label
            {
                Text = "Report created successfully.",
                Location = new Point(20, 20),
                Width = 480,
                Font = new Font(
                    SystemFonts.DefaultFont,
                    FontStyle.Bold)
            };

            var nameLabel = new Label
            {
                Text = reportName,
                Location = new Point(20, 50),
                Width = 480
            };

            var urlTextBox = new TextBox
            {
                Text = reportUrl,
                ReadOnly = true,
                Location = new Point(20, 78),
                Width = 480
            };

            var copyButton = new Button
            {
                Text = "Copy Link",
                Location = new Point(235, 125),
                Width = 80
            };

            copyButton.Click += (sender, e) =>
            {
                Clipboard.SetText(reportUrl);
            };

            var openButton = new Button
            {
                Text = "Open Report",
                Location = new Point(325, 125),
                Width = 90
            };

            openButton.Click += (sender, e) =>
            {
                Process.Start(reportUrl);
            };

            var closeButton = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Location = new Point(425, 125),
                Width = 75
            };

            Controls.Add(successLabel);
            Controls.Add(nameLabel);
            Controls.Add(urlTextBox);
            Controls.Add(copyButton);
            Controls.Add(openButton);
            Controls.Add(closeButton);

            AcceptButton = openButton;
            CancelButton = closeButton;
        }
    }
}