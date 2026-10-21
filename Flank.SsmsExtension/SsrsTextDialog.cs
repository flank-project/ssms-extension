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
        /// <summary>The query, edited by the user. Null for a stored procedure.</summary>
        public string? Sql { get; set; }
    }

    internal sealed class SsrsTextDialog : Form
    {
        private readonly ComboBox folderComboBox;
        private readonly ComboBox dataSourceComboBox;
        private readonly TextBox reportNameTextBox;
        private readonly TextBox? sqlTextBox;

        public SsrsReportOptions Options { get; private set; }

        /// <param name="sql">The query to report on, shown editable. Null for a stored procedure.</param>
        /// <param name="storedProcedure">The stored procedure to report on, shown read-only.</param>
        public SsrsTextDialog(
            string? sql,
            IReadOnlyList<string> folders,
            IReadOnlyList<string> dataSources,
            string defaultReportName = "SSMS Query Test",
            string? storedProcedure = null,
            string? defaultFolder = null,
            string? defaultDataSource = null)
        {
            Text = "Create SSRS Report";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;

            folderComboBox = AddComboBox(
                "Folder:",
                folders,
                defaultFolder,
                20);

            dataSourceComboBox = AddComboBox(
                "Shared Data Source:",
                dataSources,
                defaultDataSource,
                60);

            reportNameTextBox = AddField(
                "Report Name:",
                defaultReportName,
                100);

            int buttonsY;

            if (storedProcedure != null)
            {
                // The report calls the procedure directly (CommandType StoredProcedure).
                var procedureTextBox = AddField(
                    "Stored Procedure:",
                    storedProcedure,
                    140);

                procedureTextBox.ReadOnly = true;
                buttonsY = 185;
            }
            else
            {
                Controls.Add(new Label
                {
                    Text = "SQL:",
                    Location = new Point(20, 145),
                    Width = 130
                });

                sqlTextBox = new TextBox
                {
                    Text = sql,
                    ReadOnly = true,
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
                buttonsY = 340;
            }

            ClientSize = new Size(540, buttonsY + 40);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(350, buttonsY),
                Width = 80
            };

            var createButton = new Button
            {
                Text = "Create",
                Location = new Point(440, buttonsY),
                Width = 80
            };

            createButton.Click += CreateButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(createButton);

            AcceptButton = createButton;
            CancelButton = cancelButton;
        }

        /// <param name="preferred">Selected if it's in the list; otherwise the first item.</param>
        private ComboBox AddComboBox(
            string label,
            IReadOnlyList<string> items,
            string? preferred,
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

            int index = preferred == null ? -1 : comboBox.Items.IndexOf(preferred);

            if (index >= 0)
                comboBox.SelectedIndex = index;
            else if (comboBox.Items.Count > 0)
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
                (sqlTextBox != null && string.IsNullOrWhiteSpace(sqlTextBox.Text)))
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
                Sql = sqlTextBox?.Text
            };

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}