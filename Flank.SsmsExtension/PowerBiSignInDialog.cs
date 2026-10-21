using Flank.PowerBi;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Asks the user to sign in to Power BI: in the browser, or with a code
    /// entered in any browser (on any device). Closes with OK once signed in.
    /// </summary>
    internal sealed class PowerBiSignInDialog : Form
    {
        private readonly PowerBiSignIn signIn;
        private readonly string[]? scopes;
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly Label messageLabel;
        private readonly TextBox codeTextBox;
        private readonly Button browserButton;
        private readonly Button codeButton;
        private readonly Button openPageButton;

        public PowerBiToken? Token { get; private set; }

        /// <param name="scopes">What to ask for; Power BI by default.</param>
        public PowerBiSignInDialog(PowerBiSignIn signIn, string message, string[]? scopes = null)
        {
            this.signIn = signIn;
            this.scopes = scopes;

            Text = "Sign in to Power BI";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(500, 200);

            messageLabel = new Label
            {
                Text = message,
                Location = new Point(20, 20),
                Size = new Size(460, 60)
            };

            codeTextBox = new TextBox
            {
                ReadOnly = true,
                Location = new Point(20, 90),
                Width = 200,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 14, FontStyle.Bold),
                Visible = false
            };

            openPageButton = new Button
            {
                Text = "Open sign-in page",
                Location = new Point(235, 92),
                Width = 130,
                Visible = false
            };

            openPageButton.Click += (sender, e) => Open("https://microsoft.com/devicelogin");

            browserButton = new Button
            {
                Text = "Sign in",
                Location = new Point(185, 155),
                Width = 90
            };

            browserButton.Click += async (sender, e) =>
            {
                SetBusy("Finish signing in in your browser, then come back here.");

                try
                {
                    Finish(await signIn.SignInInteractiveAsync(Handle, cancellation.Token, scopes));
                }
                catch (Exception ex)
                {
                    Fail(ex);
                }
            };

            codeButton = new Button
            {
                Text = "Sign in with a code",
                Location = new Point(285, 155),
                Width = 120
            };

            codeButton.Click += async (sender, e) =>
            {
                SetBusy("Getting a code...");

                try
                {
                    Finish(await signIn.SignInWithDeviceCodeAsync(ShowCode, cancellation.Token, scopes));
                }
                catch (Exception ex)
                {
                    Fail(ex);
                }
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(415, 155),
                Width = 65
            };

            Controls.Add(messageLabel);
            Controls.Add(codeTextBox);
            Controls.Add(openPageButton);
            Controls.Add(browserButton);
            Controls.Add(codeButton);
            Controls.Add(cancelButton);

            AcceptButton = browserButton;
            CancelButton = cancelButton;

            FormClosed += (sender, e) => cancellation.Cancel();
        }

        private void SetBusy(string message)
        {
            messageLabel.Text = message;
            browserButton.Enabled = false;
            codeButton.Enabled = false;
        }

        // Called from MSAL's thread.
        private void ShowCode(string message)
        {
            BeginInvoke((Action)(() =>
            {
                var match = Regex.Match(message, @"code\s+(\S+)\s+to authenticate", RegexOptions.IgnoreCase);
                messageLabel.Text =
                    "Open microsoft.com/devicelogin in any browser (a phone works too), " +
                    "enter this code and sign in with your work or school account:";
                codeTextBox.Text = match.Success ? match.Groups[1].Value : message;
                codeTextBox.Visible = true;
                openPageButton.Visible = true;
            }));
        }

        private void Finish(PowerBiToken token)
        {
            Token = token;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Fail(Exception ex)
        {
            if (IsDisposed || cancellation.IsCancellationRequested)
                return;

            messageLabel.Text = "Signing in didn't work. Try again, or see the details.";
            codeTextBox.Visible = false;
            openPageButton.Visible = false;
            browserButton.Enabled = true;
            codeButton.Enabled = true;

            MessageBox.Show(this, PowerBiHelp.Explain(PowerBiStep.SignIn, ex), "Flank", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static void Open(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch (Exception)
            {
                MessageBox.Show($"Couldn't open a browser. Go to {url}", "Flank");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                cancellation.Dispose();

            base.Dispose(disposing);
        }
    }
}
