using System.Drawing;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class SsrsProgressDialog : Form
    {
        private readonly Label statusLabel;

        public SsrsProgressDialog()
        {
            Text = "Creating SSRS Report";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            ClientSize = new Size(420, 105);

            statusLabel = new Label
            {
                Text = "Analyzing query...",
                Location = new Point(20, 20),
                Width = 380
            };

            var progressBar = new ProgressBar
            {
                Location = new Point(20, 50),
                Size = new Size(380, 22),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            Controls.Add(statusLabel);
            Controls.Add(progressBar);
        }

        public void SetStatus(string status)
        {
            statusLabel.Text = status;
        }
    }
}