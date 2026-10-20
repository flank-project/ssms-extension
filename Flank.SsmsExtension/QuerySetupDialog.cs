using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>What the query setup decided: the SQL and its parameters, or the procedure made from them.</summary>
    internal sealed class QuerySetup
    {
        public QuerySetup(string sql, IReadOnlyList<SprocParameter> parameters, ProcedureSource? procedure)
        {
            Sql = sql;
            Parameters = parameters;
            Procedure = procedure;
        }

        /// <summary>The query, as edited.</summary>
        public string Sql { get; }

        /// <summary>The query's parameters (undeclared @names) with their chosen types.</summary>
        public IReadOnlyList<SprocParameter> Parameters { get; }

        /// <summary>The stored procedure created from the query, or null to run the SQL as written.</summary>
        public ProcedureSource? Procedure { get; }
    }

    /// <summary>
    /// The first step for a query from the query editor: run the SQL as
    /// written, or turn it into a stored procedure; then the SQL, where
    /// parameters are @names it doesn't DECLARE; then their types (suggested
    /// by SQL Server), and for a procedure, creating it.
    /// </summary>
    internal sealed class QuerySetupDialog : Form
    {
        private const int PageWidth = 748;

        private readonly string server;
        private readonly string database;
        private readonly Func<IDbConnection> openConnection;
        private readonly Panel[] pages;
        private int page;

        // Page 1
        private readonly RadioButton asWritten;
        private readonly RadioButton asProcedure;
        private readonly TextBox procedureName;

        // Page 2
        private readonly TextBox sqlTextBox;
        private readonly Label parametersFound;
        private readonly Timer findTimer = new Timer { Interval = 400 };

        // Page 3
        private readonly Label typesHeading;
        private readonly DataGridView typesGrid;
        private readonly Label scriptLabel;
        private readonly TextBox scriptTextBox;
        private string? typesFor;

        private readonly Button backButton;
        private readonly Button nextButton;

        public QuerySetup? Setup { get; private set; }

        /// <param name="outputName">What's being created, e.g. "workbook" or "report".</param>
        /// <param name="openConnection">Opens (or lends) a connection to the query's database.</param>
        public QuerySetupDialog(string server, string database, string sql, string outputName, Func<IDbConnection> openConnection)
        {
            this.server = server;
            this.database = database;
            this.openConnection = openConnection;

            Text = "Flank - " + database;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(780, 600);
            MinimumSize = new Size(640, 480);

            // ---- Page 1: as written, or a procedure.
            var how = NewPage();
            AddHeading(how, $"How should the {outputName} run your query?");

            asWritten = new RadioButton
            {
                Text = "Run the SQL as written",
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Location = new Point(0, 48),
                Width = PageWidth,
                Checked = true,
            };
            how.Controls.Add(asWritten);
            how.Controls.Add(Hint(
                $"The SQL goes into the {outputName}. You can add parameters on the next page.", 20, 72));

            asProcedure = new RadioButton
            {
                Text = "Turn it into a stored procedure",
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Location = new Point(0, 112),
                Width = PageWidth,
            };
            how.Controls.Add(asProcedure);
            how.Controls.Add(Hint(
                $"Flank creates the procedure (CREATE PROCEDURE), with the query's parameters, and the {outputName} " +
                "calls it. Easier to change and to grant access to later. Needs permission to create procedures.",
                20, 136));

            how.Controls.Add(new Label { Text = "Name:", Location = new Point(20, 182), Width = 60 });
            procedureName = new TextBox
            {
                Text = "dbo.",
                Location = new Point(84, 179),
                Width = 300,
                Enabled = false,
                Name = "ProcedureName",
            };
            how.Controls.Add(procedureName);

            asProcedure.CheckedChanged += (s, e) => procedureName.Enabled = asProcedure.Checked;

            // ---- Page 2: the SQL, and its parameters.
            var query = NewPage();
            AddHeading(query, "Query and parameters");
            query.Controls.Add(Hint(
                "To add a parameter, use a @name the query doesn't DECLARE, e.g. WHERE league_id = @League_Id. " +
                "You'll choose its type next.",
                0, 40));

            sqlTextBox = new TextBox
            {
                Text = sql,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Font = new Font("Consolas", 10),
                Location = new Point(0, 80),
                Size = new Size(PageWidth, query.Height - 80 - 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            query.Controls.Add(sqlTextBox);

            parametersFound = new Label
            {
                Location = new Point(0, query.Height - 32),
                Size = new Size(PageWidth, 32),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            query.Controls.Add(parametersFound);

            sqlTextBox.TextChanged += (s, e) =>
            {
                findTimer.Stop();
                findTimer.Start();
            };
            findTimer.Tick += (s, e) =>
            {
                findTimer.Stop();
                ShowParametersFound();
            };

            // ---- Page 3: types (and the procedure).
            var types = NewPage();
            typesHeading = AddHeading(types, "Parameter types");

            typesGrid = new DataGridView
            {
                Location = new Point(0, 48),
                Size = new Size(PageWidth, 180),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
            };
            typesGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Parameter", HeaderText = "Parameter", ReadOnly = true, FillWeight = 40 });
            typesGrid.Columns.Add(new DataGridViewComboBoxColumn { Name = "Type", HeaderText = "SQL type", FillWeight = 60, FlatStyle = FlatStyle.Flat });
            typesGrid.CellValueChanged += (s, e) => ShowScript();
            typesGrid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (typesGrid.IsCurrentCellDirty)
                    typesGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            types.Controls.Add(typesGrid);

            scriptLabel = new Label
            {
                Text = "Flank will run this to create the procedure:",
                Location = new Point(0, 240),
                Width = PageWidth,
            };
            types.Controls.Add(scriptLabel);

            scriptTextBox = new TextBox
            {
                ReadOnly = true,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9),
                Location = new Point(0, 262),
                Size = new Size(PageWidth, types.Height - 262),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = SystemColors.Window,
            };
            types.Controls.Add(scriptTextBox);

            pages = new[] { how, query, types };

            backButton = new Button
            {
                Text = "< Back",
                Location = new Point(ClientSize.Width - 316, ClientSize.Height - 40),
                Width = 90,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            backButton.Click += (s, e) => ShowPage(page == 2 ? 1 : 0);

            nextButton = new Button
            {
                Location = new Point(ClientSize.Width - 216, ClientSize.Height - 40),
                Width = 110,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            nextButton.Click += (s, e) => Next();

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(ClientSize.Width - 96, ClientSize.Height - 40),
                Width = 80,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };

            Controls.AddRange(new Control[] { backButton, nextButton, cancelButton });
            CancelButton = cancelButton;

            ShowPage(0);
        }

        private Panel NewPage()
        {
            var panel = new Panel
            {
                Location = new Point(16, 12),
                Size = new Size(PageWidth, ClientSize.Height - 12 - 56),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false,
            };
            Controls.Add(panel);
            return panel;
        }

        private static Label AddHeading(Panel panel, string text)
        {
            var heading = new Label
            {
                Text = text,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 14, FontStyle.Bold),
                Location = new Point(0, 0),
                Size = new Size(PageWidth, 32),
            };
            panel.Controls.Add(heading);
            return heading;
        }

        private static Label Hint(string text, int x, int y) =>
            new Label
            {
                Text = text,
                ForeColor = SystemColors.GrayText,
                Location = new Point(x, y),
                Size = new Size(PageWidth - x, 34),
            };

        private IReadOnlyList<string> CurrentParameters => QueryParameters.Find(sqlTextBox.Text);

        private bool HasTypesPage => asProcedure.Checked || CurrentParameters.Count > 0;

        private void ShowPage(int index)
        {
            page = index;

            for (int i = 0; i < pages.Length; i++)
                pages[i].Visible = i == index;

            backButton.Enabled = index > 0;

            bool last = index == 2 || (index == 1 && !HasTypesPage);
            nextButton.Text = !last ? "Next >" : asProcedure.Checked ? "Create procedure" : "Continue";
            AcceptButton = index == 1 ? null : nextButton; // Return types a new line in the SQL

            if (index == 1)
            {
                ShowParametersFound();
                sqlTextBox.Select();
            }
            else if (index == 2)
            {
                ShowTypes();
            }
        }

        private void ShowParametersFound()
        {
            string? problem = QueryParameters.ParseProblem(sqlTextBox.Text);
            var found = CurrentParameters;

            parametersFound.ForeColor = problem != null ? Color.Firebrick : SystemColors.ControlText;
            parametersFound.Text =
                problem != null ? "The query doesn't parse. " + problem :
                found.Count == 0 ? "No parameters." :
                "Parameters: " + string.Join(", ", found);

            if (page == 1)
                nextButton.Text = HasTypesPage ? "Next >" : "Continue";
        }

        /// <summary>One row per parameter, with SQL Server's suggested type (kept if the user already chose one).</summary>
        private void ShowTypes()
        {
            var names = CurrentParameters;

            var chosen = typesGrid.Rows.Cast<DataGridViewRow>()
                .ToDictionary(r => (string)r.Cells["Parameter"].Value, r => r.Cells["Type"].Value as string, StringComparer.OrdinalIgnoreCase);

            IReadOnlyDictionary<string, string> suggested = new Dictionary<string, string>();

            if (names.Count > 0 && typesFor != sqlTextBox.Text)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;

                    using (var connection = openConnection())
                    {
                        suggested = QueryParameters.SuggestTypes(sqlTextBox.Text, connection);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Flank: no type suggestions: {ex.Message}");
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

                typesFor = sqlTextBox.Text;
            }

            typesGrid.Rows.Clear();

            foreach (string name in names)
            {
                string type =
                    chosen.TryGetValue(name, out string? previous) && previous != null ? previous :
                    suggested.TryGetValue(name, out string? suggestion) ? suggestion :
                    QueryParameters.DefaultType;

                int row = typesGrid.Rows.Add(name, null);
                var cell = (DataGridViewComboBoxCell)typesGrid.Rows[row].Cells["Type"];
                cell.Items.AddRange(QueryParameters.CommonTypes.Cast<object>().ToArray());

                if (!cell.Items.Contains(type))
                    cell.Items.Insert(0, type);

                cell.Value = type;
            }

            typesGrid.Visible = names.Count > 0;
            typesHeading.Text = names.Count == 0 ? "Create the procedure" : "Parameter types";

            bool procedure = asProcedure.Checked;
            scriptLabel.Visible = procedure;
            scriptTextBox.Visible = procedure;
            scriptLabel.Top = names.Count > 0 ? 240 : 48;
            scriptTextBox.Top = scriptLabel.Top + 22;
            scriptTextBox.Height = pages[2].Height - scriptTextBox.Top;

            ShowScript();
        }

        private IReadOnlyList<SprocParameter>? ChosenParameters()
        {
            var parameters = new List<SprocParameter>();

            foreach (DataGridViewRow row in typesGrid.Rows)
            {
                string name = (string)row.Cells["Parameter"].Value;
                var parameter = QueryParameters.Parameter(name, row.Cells["Type"].Value as string ?? "");

                if (parameter == null)
                    return null;

                parameters.Add(parameter);
            }

            return parameters;
        }

        private void ShowScript()
        {
            if (!asProcedure.Checked)
                return;

            var parameters = ChosenParameters() ?? new List<SprocParameter>();
            scriptTextBox.Text = QueryParameters
                .CreateProcedureScript(NormalizeProcedureName(procedureName.Text) ?? procedureName.Text, parameters, sqlTextBox.Text)
                .Replace("\n", Environment.NewLine);
        }

        private void Next()
        {
            if (page == 0)
            {
                if (asProcedure.Checked && NormalizeProcedureName(procedureName.Text) == null)
                {
                    Problem("Enter the new procedure's name, e.g. dbo.monthly_sales.");
                    procedureName.Select();
                    return;
                }

                ShowPage(1);
                return;
            }

            if (page == 1)
            {
                findTimer.Stop();
                string? problem = QueryParameters.ParseProblem(sqlTextBox.Text);

                if (string.IsNullOrWhiteSpace(sqlTextBox.Text))
                {
                    Problem("The query is empty.");
                    return;
                }

                if (problem != null)
                {
                    Problem("The query doesn't parse, so Flank can't find its parameters.\n\n" + problem);
                    return;
                }

                if (HasTypesPage)
                {
                    ShowPage(2);
                    return;
                }

                Finish(new List<SprocParameter>());
                return;
            }

            typesGrid.EndEdit();
            var parameters = ChosenParameters();

            if (parameters == null)
            {
                Problem("Choose a SQL type for each parameter.");
                return;
            }

            Finish(parameters);
        }

        private void Finish(IReadOnlyList<SprocParameter> parameters)
        {
            string sql = sqlTextBox.Text;

            if (!asProcedure.Checked)
            {
                Setup = new QuerySetup(sql, parameters, null);
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            string name = NormalizeProcedureName(procedureName.Text)!;
            string script = QueryParameters.CreateProcedureScript(name, parameters, sql);

            try
            {
                Cursor = Cursors.WaitCursor;

                using (var connection = openConnection())
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = script;
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Problem(
                    "Flank couldn't create the procedure.\n\n" + ex.Message +
                    (ex.Message.IndexOf("already", StringComparison.OrdinalIgnoreCase) >= 0
                        ? "\n\nChoose another name on the first page."
                        : ""));
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            var parts = name.Split(new[] { "].[" }, StringSplitOptions.None);
            string schema = parts[0].TrimStart('[').Replace("]]", "]");
            string procedure = parts[1].TrimEnd(']').Replace("]]", "]");

            Setup = new QuerySetup(sql, parameters, new ProcedureSource(server, database, schema, procedure, openConnection));
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Problem(string message) =>
            MessageBox.Show(this, message, "Flank", MessageBoxButtons.OK, MessageBoxIcon.Information);

        /// <summary>"monthly", "dbo.monthly" or "[dbo].[monthly]" -> "[dbo].[monthly]"; null if invalid.</summary>
        private static string? NormalizeProcedureName(string text)
        {
            var parts = text.Trim().Replace("[", "").Replace("]", "").Split('.');

            if (parts.Length > 2 || parts.Any(p => string.IsNullOrWhiteSpace(p)))
                return null;

            string schema = parts.Length == 2 ? parts[0].Trim() : "dbo";
            string name = parts[parts.Length - 1].Trim();

            return AccessScript.QuoteName(schema) + "." + AccessScript.QuoteName(name);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                findTimer.Dispose();

            base.Dispose(disposing);
        }
    }
}
