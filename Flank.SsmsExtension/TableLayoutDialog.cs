using Flank.Ssrs;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Where a workbook has people enter a table-valued parameter's rows: in
    /// the parameter list (a set number of rows to start with), or on a sheet
    /// of its own.
    /// </summary>
    internal sealed class TableLayoutDialog : Form
    {
        private readonly ParameterSettings setting;
        private readonly RadioButton inline;
        private readonly NumericUpDown rows;
        private readonly RadioButton sheet;

        public TableLayoutDialog(ParameterSettings setting)
        {
            this.setting = setting;

            Text = "Flank - " + setting.DisplayName;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new Size(560, 250);

            Controls.Add(new Label
            {
                Text = "Where people enter the rows",
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 11f, FontStyle.Bold),
                Location = new Point(14, 12),
                AutoSize = true,
            });

            Controls.Add(new Label
            {
                Text = $"{setting.SqlName} {setting.Parameter.TableType}: " +
                       string.Join(", ", setting.Parameter.Columns.Select(c => c.Name)),
                ForeColor = SystemColors.GrayText,
                Location = new Point(16, 40),
                Size = new Size(528, 20),
                AutoEllipsis = true,
            });

            inline = new RadioButton
            {
                Text = "In the parameter list, starting with",
                Location = new Point(16, 72),
                AutoSize = true,
                Checked = setting.TableLayout == TableLayout.Inline,
            };
            rows = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 100,
                Value = Math.Max(1, Math.Min(100, setting.TableRows)),
                Location = new Point(250, 70),
                Width = 56,
            };
            Controls.Add(new Label { Text = "rows", Location = new Point(312, 74), AutoSize = true });
            Controls.Add(new Label
            {
                Text = "For a few rows, about the same number each time. People can insert more rows.",
                ForeColor = SystemColors.GrayText,
                Location = new Point(34, 98),
                Size = new Size(510, 20),
            });

            sheet = new RadioButton
            {
                Text = "On a sheet of its own, linked from the parameter list",
                Location = new Point(16, 130),
                AutoSize = true,
                Checked = setting.TableLayout == TableLayout.Sheet,
            };
            Controls.Add(new Label
            {
                Text = "For pasting in many rows, or when the number of rows varies a lot.",
                ForeColor = SystemColors.GrayText,
                Location = new Point(34, 156),
                Size = new Size(510, 20),
            });

            inline.CheckedChanged += (s, e) => rows.Enabled = inline.Checked;
            rows.ValueChanged += (s, e) => inline.Checked = true;
            rows.Enabled = inline.Checked;

            var ok = new Button { Text = "OK", Location = new Point(ClientSize.Width - 206, ClientSize.Height - 40), Width = 90 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(ClientSize.Width - 106, ClientSize.Height - 40), Width = 90 };
            ok.Click += (s, e) =>
            {
                this.setting.TableLayout = inline.Checked ? TableLayout.Inline : TableLayout.Sheet;
                this.setting.TableRows = (int)rows.Value;
                DialogResult = DialogResult.OK;
                Close();
            };

            Controls.AddRange(new Control[] { inline, rows, sheet, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <summary>"In the list, 3 rows" or "Own sheet", for the review page.</summary>
        public static string Summary(ParameterSettings setting) =>
            setting.TableLayout == TableLayout.Sheet
                ? "Rows on their own sheet"
                : $"Rows in the list, {setting.TableRows} to start";
    }
}
