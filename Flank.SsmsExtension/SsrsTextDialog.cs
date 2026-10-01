using System;
using System.Drawing;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class SsrsReportOptions
    {
        public string ReportServerUrl { get; set; }
        public string ReportPortalUrl { get; set; }
        public string Folder { get; set; }
        public string DataSource { get; set; }
        public string ReportName { get; set; }
        public string Sql { get; set; }
    }

    internal sealed class SsrsTextDialog : Form
    {
        private readonly TextBox reportServerTextBox;
        private readonly TextBox reportPortalTextBox;
        private readonly TextBox folderTextBox;
        private readonly TextBox dataSourceTextBox;
        private readonly TextBox reportNameTextBox;

        public SsrsReportOptions Options { get; private set; }
        private readonly TextBox sqlTextBox;

        public SsrsTextDialog(string sql)
        {
            Text = "Create SSRS Report";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(540, 460);

            reportServerTextBox = AddField(
                "Report Server URL:",
                "http://localhost/ReportServer",
                20);

            reportPortalTextBox = AddField(
                "Report Portal URL:",
                "http://localhost/Reports",
                60);

            folderTextBox = AddField(
                "Folder:",
                "/TestFolder",
                100);

            dataSourceTextBox = AddField(
                "Shared Data Source:",
                "/HardcodedTest",
                140);

            reportNameTextBox = AddField(
                "Report Name:",
                "SSMS Query Test",
                180);

            sqlTextBox = new TextBox
            {
                Text = sql,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Font = new Font("Consolas", 9),
                Location = new Point(20, 225),
                Size = new Size(500, 180)
            };

            Controls.Add(new Label
            {
                Text = "SQL:",
                Location = new Point(20, 205),
                Width = 130
            });

            Controls.Add(sqlTextBox);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(350, 420),
                Width = 80
            };

            var createButton = new Button
            {
                Text = "Create",
                Location = new Point(440, 420),
                Width = 80
            };

            createButton.Click += CreateButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(createButton);

            AcceptButton = createButton;
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

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(reportServerTextBox.Text) ||
                string.IsNullOrWhiteSpace(reportPortalTextBox.Text) ||
                string.IsNullOrWhiteSpace(folderTextBox.Text) ||
                string.IsNullOrWhiteSpace(dataSourceTextBox.Text) ||
                string.IsNullOrWhiteSpace(reportNameTextBox.Text) ||
                string.IsNullOrWhiteSpace(sqlTextBox.Text))
            {
                MessageBox.Show(
                    "All fields are required.",
                    "Flank",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            Options = new SsrsReportOptions
            {
                ReportServerUrl = reportServerTextBox.Text.Trim(),
                ReportPortalUrl = reportPortalTextBox.Text.Trim(),
                Folder = folderTextBox.Text.Trim(),
                DataSource = dataSourceTextBox.Text.Trim(),
                ReportName = reportNameTextBox.Text.Trim(),
                Sql = sqlTextBox.Text
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}