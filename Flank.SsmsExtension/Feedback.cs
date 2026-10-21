using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Ways to reach the developer, from the Feedback menu item.
    /// </summary>
    internal static class Feedback
    {
        public const string Email = "anguspmitchell@gmail.com";
        public const string DiscussionsUrl = "https://github.com/flank-project/ssms-extension/discussions";

        public static void ShowDialog()
        {
            using (var dialog = new FeedbackDialog(
                "Feedback",
                "Questions, ideas or problems? I'd love to hear how you use Flank."))
            {
                dialog.ShowDialog();
            }
        }

        private sealed class FeedbackDialog : Form
        {
            public FeedbackDialog(string title, string message)
            {
                Text = title;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                ClientSize = new Size(480, 175);

                var messageLabel = new Label
                {
                    Text = message,
                    Location = new Point(20, 20),
                    Size = new Size(440, 50)
                };

                var emailTextBox = new TextBox
                {
                    Text = Email,
                    ReadOnly = true,
                    Location = new Point(20, 78),
                    Width = 440
                };

                var emailButton = new Button
                {
                    Text = "Copy Email",
                    Location = new Point(135, 125),
                    Width = 85
                };

                emailButton.Click += (sender, e) =>
                {
                    Clipboard.SetText(Email);
                    emailButton.Text = "Copied";
                };

                var discussionsButton = new Button
                {
                    Text = "Open GitHub Discussions",
                    Location = new Point(230, 125),
                    Width = 155
                };

                discussionsButton.Click += (sender, e) =>
                    Open(DiscussionsUrl,
                        $"Couldn't open a browser. The address is {DiscussionsUrl}");

                var closeButton = new Button
                {
                    Text = "Close",
                    DialogResult = DialogResult.OK,
                    Location = new Point(395, 125),
                    Width = 65
                };

                Controls.Add(messageLabel);
                Controls.Add(emailTextBox);
                Controls.Add(emailButton);
                Controls.Add(discussionsButton);
                Controls.Add(closeButton);

                AcceptButton = emailButton;
                CancelButton = closeButton;
            }

            private static void Open(string url, string failureMessage)
            {
                try
                {
                    Process.Start(url);
                }
                catch (Exception)
                {
                    MessageBox.Show(failureMessage, "Flank");
                }
            }
        }
    }
}
