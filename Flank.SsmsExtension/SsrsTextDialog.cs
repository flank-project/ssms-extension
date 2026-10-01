using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class SsrsReportOptions
    {
        public string Folder { get; set; }
        public string DataSource { get; set; }
        public string ReportName { get; set; }
        public string Sql { get; set; }
    }

    internal sealed class SsrsTextDialog : Form
    {
        private readonly ComboBox folderComboBox;
        private readonly ComboBox dataSourceComboBox;
        private readonly TextBox reportNameTextBox;
        private readonly TextBox sqlTextBox;

        public SsrsReportOptions Options { get; private set; }

        public SsrsTextDialog(
            string sql,
            IReadOnlyList<string> folders,
            IReadOnlyList<string> dataSources)
        {
            Text = "Create SSRS Report";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(540, 380);

            folderComboBox = AddComboBox(
                "Folder:",
                folders,
                20);

            dataSourceComboBox = AddComboBox(
                "Shared Data Source:",
                dataSources,
                60);

            reportNameTextBox = AddField(
                "Report Name:",
                "SSMS Query Test",
                100);

            Controls.Add(new Label
            {
                Text = "SQL:",
                Location = new Point(20, 145),
                Width = 130
            });

            sqlTextBox = new TextBox
            {
                Text = sql,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Font = new Font("Consolas", 9),
                Location = new Point(20, 165),
                Size = new Size(500, 160)
            };

            Controls.Add(sqlTextBox);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(350, 340),
                Width = 80
            };

            var createButton = new Button
            {
                Text = "Create",
                Location = new Point(440, 340),
                Width = 80
            };

            createButton.Click += CreateButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(createButton);

            AcceptButton = createButton;
            CancelButton = cancelButton;
        }

        private ComboBox AddComboBox(
            string label,
            IReadOnlyList<string> items,
            int y)
        {
            var labelControl = new Label
            {
                Text = label,
                Location = new Point(20, y + 4),
                Width = 130
            };

            var comboBox = new ComboBox
            {
                Location = new Point(155, y),
                Width = 365,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            comboBox.Items.AddRange(items.ToArray());

            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0;

            Controls.Add(labelControl);
            Controls.Add(comboBox);

            return comboBox;
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
            if (folderComboBox.SelectedItem == null ||
                dataSourceComboBox.SelectedItem == null ||
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
                Folder = folderComboBox.SelectedItem.ToString(),
                DataSource = dataSourceComboBox.SelectedItem.ToString(),
                ReportName = reportNameTextBox.Text.Trim(),
                Sql = sqlTextBox.Text
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}