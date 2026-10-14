using Flank.Excel;
using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Creating a refreshable workbook, from a query or a stored procedure:
    /// 1. who will refresh it (skip, or an existing / new database user),
    /// 2. what they can do (run the procedure; for a query, turn it into a
    ///    procedure or read its tables): the script shown runs on Run SQL,
    /// 3. run it as that user to check and preview, then create the workbook
    ///    with those values and rows.
    /// </summary>
    internal sealed class ExcelWorkbookWizard : Form
    {
        private const int PageWidth = 748;

        private readonly string server;
        private readonly string database;
        private readonly bool isAzureSqlDatabase;
        private readonly string? procedure;
        private readonly IReadOnlyList<SprocParameter> procedureParameters;
        private readonly string? query;
        private readonly IReadOnlyList<string> queryTables;
        private readonly string? queryTablesProblem;
        private readonly Action<Action<IDbConnection>> useConnection;
        private readonly List<AccessSignIn?> signInOptions;

        private readonly Panel[] pages;
        private int page;

        // Page 1: who.
        private readonly List<RadioButton> signInRadios = new List<RadioButton>();
        private readonly GroupBox userGroup;
        private readonly RadioButton existingUser;
        private readonly ComboBox existingUsers;
        private readonly RadioButton newUser;
        private readonly TextBox newUserName;
        private readonly Label passwordLabel;
        private readonly TextBox password;
        private readonly Button newPassword;

        // Page 2: what.
        private readonly RadioButton makeProcedure;
        private readonly TextBox procedureName;
        private readonly RadioButton readTables;
        private readonly TextBox script;
        private readonly Label instructions;
        private string? appliedScript;

        // Page 3: run.
        private readonly Panel runPage;
        private readonly Label runHeading;
        private readonly CheckBox runAsMe;
        private readonly Label runNote;
        private string? me;
        private PreviewPanel? preview;

        private readonly Button backButton;
        private readonly Button nextButton;

        /// <param name="procedure">[schema].[name] when creating from a stored procedure.</param>
        /// <param name="query">The query, when creating from the query editor.</param>
        /// <param name="queryTables">Tables/views the query reads, resolved as [schema].[name].</param>
        /// <param name="queryTablesProblem">Why <paramref name="queryTables"/> couldn't be found, if not.</param>
        /// <param name="useConnection">
        /// Calls the action with a connection to the database. It may be
        /// borrowed (the query window's), so it's never closed here.
        /// </param>
        public ExcelWorkbookWizard(
            string server,
            string database,
            bool isAzureSqlDatabase,
            string? procedure,
            IReadOnlyList<SprocParameter> procedureParameters,
            string? query,
            IReadOnlyList<string> queryTables,
            string? queryTablesProblem,
            Action<Action<IDbConnection>> useConnection)
        {
            this.server = server;
            this.database = database;
            this.isAzureSqlDatabase = isAzureSqlDatabase;
            this.procedure = procedure;
            this.procedureParameters = procedureParameters;
            this.query = query;
            this.queryTables = queryTables;
            this.queryTablesProblem = queryTablesProblem;
            this.useConnection = useConnection;

            Text = "Create Excel workbook - " + (procedure ?? database);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(780, 640);
            MinimumSize = new Size(640, 560);

            // ---- Page 1: who will refresh it.
            var who = NewPage();
            AddHeading(who, "Who will refresh this workbook?", 0);

            // Azure SQL Database has no Windows logins; SQL Server users don't
            // normally have Entra accounts in the database.
            signInOptions = isAzureSqlDatabase
                ? new List<AccessSignIn?> { null, AccessSignIn.SqlLogin, AccessSignIn.Entra }
                : new List<AccessSignIn?> { null, AccessSignIn.SqlLogin, AccessSignIn.Windows };

            int y = 34;
            foreach (var option in signInOptions)
            {
                var radio = new RadioButton
                {
                    Text = Describe(option),
                    Location = new Point(0, y),
                    Size = new Size(PageWidth, 24),
                    Checked = option == null,
                };
                radio.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) SignInChanged(); };
                signInRadios.Add(radio);
                who.Controls.Add(radio);
                y += 26;
            }

            userGroup = new GroupBox
            {
                Text = "Which user?",
                Location = new Point(0, y + 10),
                Size = new Size(PageWidth, 150),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            who.Controls.Add(userGroup);

            existingUser = new RadioButton { Text = "An existing database user:", Location = new Point(14, 26), Size = new Size(200, 24) };
            existingUsers = new ComboBox { Location = new Point(220, 26), Width = 360, DropDownStyle = ComboBoxStyle.DropDownList };
            newUser = new RadioButton { Text = "A new user named:", Location = new Point(14, 62), Size = new Size(200, 24) };
            newUserName = new TextBox { Location = new Point(220, 62), Width = 360, Name = "NewUserName" };
            passwordLabel = new Label { Text = "Password:", Location = new Point(34, 100), Size = new Size(180, 20) };
            password = new TextBox { Location = new Point(220, 98), Width = 260, Text = AccessScript.GeneratePassword(), Name = "NewUserPassword" };
            newPassword = new Button { Text = "New password", Location = new Point(490, 97), Width = 90 };

            existingUser.CheckedChanged += (s, e) => UpdateButtons();
            newUser.CheckedChanged += (s, e) => UpdateButtons();
            existingUsers.SelectedIndexChanged += (s, e) => { existingUser.Checked = true; UpdateButtons(); };
            newUserName.TextChanged += (s, e) => { if (newUserName.Text.Length > 0) newUser.Checked = true; UpdateButtons(); };
            newPassword.Click += (s, e) => password.Text = AccessScript.GeneratePassword();

            userGroup.Controls.AddRange(new Control[] { existingUser, existingUsers, newUser, newUserName, passwordLabel, password, newPassword });

            // ---- Page 2: what they can do.
            var what = NewPage();
            AddHeading(what, "What can they do?", 0);
            y = 34;

            if (procedure != null)
            {
                what.Controls.Add(new Label
                {
                    Text = $"Run {procedure} (GRANT EXECUTE). The procedure decides what data they see.",
                    Location = new Point(0, y),
                    Size = new Size(PageWidth, 22),
                });
                y += 30;
            }

            makeProcedure = new RadioButton
            {
                Text = "Turn the query into a stored procedure, and let them run it:",
                Location = new Point(0, y),
                Size = new Size(420, 24),
                Visible = procedure == null,
                Checked = procedure == null,
            };
            procedureName = new TextBox
            {
                Location = new Point(425, y),
                Width = 300,
                Text = "[dbo].[flank_report]",
                Visible = procedure == null,
                Name = "NewProcedureName",
            };
            readTables = new RadioButton
            {
                Text = queryTables.Count > 0
                    ? "Let them read the tables the query uses: " + string.Join(", ", queryTables)
                    : "Let them read the tables the query uses (" + (queryTablesProblem ?? "none found") + ")",
                Location = new Point(0, y + 28),
                Size = new Size(PageWidth, 24),
                AutoEllipsis = true,
                Visible = procedure == null,
                Enabled = queryTables.Count > 0,
            };
            makeProcedure.CheckedChanged += (s, e) => AccessChanged();
            readTables.CheckedChanged += (s, e) => AccessChanged();
            procedureName.TextChanged += (s, e) => AccessChanged();
            what.Controls.AddRange(new Control[] { makeProcedure, procedureName, readTables });
            if (procedure == null) y += 62;

            what.Controls.Add(new Label { Text = "Script:", Location = new Point(0, y), AutoSize = true });
            script = new TextBox
            {
                Location = new Point(0, y + 20),
                Size = new Size(PageWidth, 200),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Name = "AccessScript",
            };
            what.Controls.Add(script);

            instructions = new Label
            {
                Location = new Point(0, y + 230),
                Size = new Size(PageWidth, 70),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            what.Controls.Add(instructions);

            // ---- Page 3: run as them.
            runPage = NewPage();
            runHeading = AddHeading(runPage, "", 0);
            runAsMe = new CheckBox
            {
                Text = "Run as me instead",
                Location = new Point(0, 30),
                AutoSize = true,
                Visible = false,
            };
            runAsMe.CheckedChanged += (s, e) => UpdateRunHeading();
            runNote = new Label
            {
                Location = new Point(0, 30),
                Size = new Size(PageWidth, 22),
                AutoEllipsis = true,
                ForeColor = SystemColors.GrayText,
                Visible = false,
            };
            runPage.Controls.AddRange(new Control[] { runAsMe, runNote });

            pages = new[] { who, what, runPage };

            // ---- Navigation.
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90, Location = new Point(ClientSize.Width - 106, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            nextButton = new Button { Text = "Next >", Width = 120, Location = new Point(ClientSize.Width - 236, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            backButton = new Button { Text = "< Back", Width = 90, Location = new Point(ClientSize.Width - 336, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            nextButton.Click += (s, e) => Next();
            backButton.Click += (s, e) => Back();
            Controls.AddRange(new Control[] { backButton, nextButton, cancel });
            CancelButton = cancel;

            ShowPage(0);
            SignInChanged();
        }

        /// <summary>The workbook to create, once the wizard finishes.</summary>
        public WorkbookRequest? Request { get; private set; }

        public string DefaultFileName =>
            (procedure != null ? StripBrackets(procedure).Split('.').Last() : "report") + ".xlsx";

        private AccessSignIn? SelectedSignIn => signInOptions[signInRadios.FindIndex(r => r.Checked)];

        private string? UserName =>
            SelectedSignIn == null ? null
            : existingUser.Checked ? existingUsers.SelectedItem as string
            : newUserName.Text.Trim();

        /// <summary>The procedure the workbook calls: the given one, or the query turned into one.</summary>
        private string? EffectiveProcedure =>
            procedure ?? (SelectedSignIn != null && makeProcedure.Checked ? NormalizeProcedureName(procedureName.Text) : null);

        // ---------------------------------------------------------------- pages

        private Panel NewPage()
        {
            var panel = new Panel
            {
                Location = new Point(16, 16),
                Size = new Size(PageWidth, ClientSize.Height - 16 - 56),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false,
            };
            Controls.Add(panel);
            return panel;
        }

        private static Label AddHeading(Panel panel, string text, int y)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(0, y),
                Size = new Size(PageWidth, 26),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 11f, FontStyle.Bold),
            };
            panel.Controls.Add(label);
            return label;
        }

        private void ShowPage(int index)
        {
            page = index;
            for (int i = 0; i < pages.Length; i++)
                pages[i].Visible = i == index;

            if (index == 1)
                AccessChanged();
            if (index == 2)
                BuildRunPage();

            UpdateButtons();

            if (index == 2)
                preview?.FocusFirstInput();
        }

        private void Next()
        {
            if (page == 0)
            {
                // Skipping: nothing to set up, straight to running as you.
                ShowPage(SelectedSignIn == null ? 2 : 1);
            }
            else if (page == 1)
            {
                if (ApplyAccess())
                    ShowPage(2);
            }
            else
            {
                Finish();
            }
        }

        private void Back()
        {
            if (page == 2)
                ShowPage(SelectedSignIn == null ? 0 : 1);
            else if (page == 1)
                ShowPage(0);
        }

        private void UpdateButtons()
        {
            // Only a new SQL user needs a password.
            passwordLabel.Visible = password.Visible = newPassword.Visible =
                SelectedSignIn == AccessSignIn.SqlLogin && newUser.Checked;

            backButton.Enabled = page > 0;
            nextButton.Text = page == 1 ? "Run SQL" : page == 2 ? "Create workbook" : "Next >";

            nextButton.Enabled =
                page == 0 ? SelectedSignIn == null || !string.IsNullOrWhiteSpace(UserName) &&
                            (existingUser.Checked || SelectedSignIn != AccessSignIn.SqlLogin || password.Text.Length > 0)
                : page == 1 ? TryBuildAccess(out _)
                : preview != null && preview.HasResults;

            AcceptButton = page == 2 && preview != null && !preview.HasResults ? preview.RunButton : nextButton;
        }

        // --------------------------------------------------------------- page 1

        private static string Describe(AccessSignIn? signIn)
        {
            switch (signIn)
            {
                case null: return "Skip this if end user already has access";
                case AccessSignIn.Entra: return "Someone with a Microsoft Entra account";
                case AccessSignIn.Windows: return "Someone with a Windows / Active Directory account";
                default: return "Someone with a SQL Server username and password";
            }
        }

        private void SignInChanged()
        {
            var signIn = SelectedSignIn;
            userGroup.Enabled = signIn != null;
            newUser.Text = signIn == AccessSignIn.Entra ? "A new user (email / UPN):"
                         : signIn == AccessSignIn.Windows ? @"A new user (DOMAIN\user):"
                         : "A new user named:";

            existingUsers.Items.Clear();

            if (signIn != null)
            {
                try
                {
                    IReadOnlyList<string> users = Array.Empty<string>();
                    useConnection(connection => users = AccessDiscovery.ListUsers(connection, signIn.Value));
                    existingUsers.Items.AddRange(users.Cast<object>().ToArray());
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Flank: could not list users: {ex.Message}");
                }
            }

            bool any = existingUsers.Items.Count > 0;
            existingUser.Enabled = existingUsers.Enabled = any;
            if (any)
            {
                existingUsers.SelectedIndex = 0;
                existingUser.Checked = true;
            }
            else
            {
                newUser.Checked = true;
            }

            UpdateButtons();
        }

        // --------------------------------------------------------------- page 2

        private bool TryBuildAccess(out AccessRequest request)
        {
            string? effective = EffectiveProcedure;

            request = new AccessRequest
            {
                SignIn = SelectedSignIn ?? AccessSignIn.SqlLogin,
                Name = UserName ?? "",
                ExistingUser = existingUser.Checked,
                Password = password.Text,
                Database = database,
                IsAzureSqlDatabase = isAzureSqlDatabase,
                Scope = effective != null ? AccessScope.Procedure : AccessScope.Objects,
                Procedure = effective ?? "",
                ProcedureBody = procedure == null && effective != null ? query : null,
                Objects = queryTables,
            };

            return request.Name.Length > 0 &&
                   (effective != null || queryTables.Count > 0) &&
                   (procedure != null || !makeProcedure.Checked || effective != null);
        }

        private void AccessChanged()
        {
            if (TryBuildAccess(out var request))
            {
                script.Text = AccessScript.ToText(request);
            }
            else
            {
                script.Text = makeProcedure.Checked && EffectiveProcedure == null
                    ? "-- Enter a procedure name, e.g. [dbo].[monthly_sales]."
                    : "-- Nothing to grant.";
            }

            string who = UserName ?? "the user";
            switch (SelectedSignIn)
            {
                case AccessSignIn.SqlLogin:
                    instructions.Text = existingUser.Checked
                        ? $"Send them the workbook. When Excel asks them to sign in, they choose Database and use {who}'s username and password."
                        : $"Send them the workbook, plus: server {server}, database {database}, username {who} and the password from the previous step.\n" +
                          "When Excel asks them to sign in, they choose Database and enter that username and password.";
                    break;
                case AccessSignIn.Entra:
                    instructions.Text =
                        $"{who} must be in your Microsoft Entra tenant. Creating Entra users needs an Entra-authenticated connection.\n" +
                        "When Excel asks them to sign in, they choose Microsoft account and use that Entra account.";
                    break;
                case AccessSignIn.Windows:
                    instructions.Text =
                        $"Their computer must be able to reach {server} (e.g. on the network or VPN).\n" +
                        "When Excel asks them to sign in, they choose Windows and use their current credentials.";
                    break;
            }

            UpdateButtons();
        }

        /// <summary>Runs the script shown; on failure says why and stays on this page.</summary>
        private bool ApplyAccess()
        {
            if (!TryBuildAccess(out var request))
                return false;

            // Back then Run SQL again: the same script already ran (and
            // CREATE PROCEDURE would fail the second time).
            string text = AccessScript.ToText(request);
            if (text == appliedScript)
                return true;

            Cursor = Cursors.WaitCursor;
            try
            {
                useConnection(connection => AccessScript.Run(connection, request));
                appliedScript = text;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "The script didn't run:\n\n" + ex.Message, "Flank", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // --------------------------------------------------------------- page 3

        private void BuildRunPage()
        {
            if (preview != null)
            {
                runPage.Controls.Remove(preview);
                preview.Dispose();
            }

            string? effective = EffectiveProcedure;
            var inputs = effective == procedure && procedure != null
                ? procedureParameters.Where(p => !p.IsOutput).ToList()
                : new List<SprocParameter>();
            var outputs = effective == procedure && procedure != null
                ? procedureParameters.Where(p => p.IsOutput).ToList()
                : new List<SprocParameter>();

            if (me == null)
            {
                try
                {
                    useConnection(connection => me = AccessTester.CurrentLogin(connection));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Flank: could not read the current login: {ex.Message}");
                }
            }

            runAsMe.Text = $"Run as you ({Me}) instead";
            runAsMe.Visible = false;
            runAsMe.Checked = !CanRunAsUser;

            runNote.Text = SelectedSignIn == AccessSignIn.Entra
                ? $"Flank can't run as a Microsoft Entra user, so this runs as you. {UserName} sees what their grants allow."
                : "";
            runNote.Visible = runNote.Text.Length > 0;

            preview = new PreviewPanel(inputs, new Size(PageWidth, runPage.Height - 60))
            {
                Location = new Point(0, 58),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };

            preview.Runner = values =>
            {
                var named = inputs.Select((p, i) => new KeyValuePair<string, object?>(p.Name, values[i])).ToList();
                string? user = runAsMe.Checked ? null : UserName;
                IReadOnlyList<ResultSet> results = Array.Empty<ResultSet>();

                useConnection(connection =>
                {
                    Func<IReadOnlyList<ResultSet>> run = effective != null
                        ? () => StoredProcedureRunner.Run(connection, effective, named, PreviewPanel.MaxRows, outputs)
                        : () => StoredProcedureRunner.RunQuery(connection, query!, PreviewPanel.MaxRows);

                    results = user == null ? run() : AccessTester.RunAs(connection, user, run);
                });

                return results;
            };

            preview.ResultsChanged += (s, e) => UpdateButtons();
            preview.RunFailed += (s, ex) =>
            {
                // Impersonating needs IMPERSONATE permission (e.g. db_owner),
                // which granting access doesn't.
                if (!runAsMe.Checked && ex.Message.IndexOf("Cannot execute as", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    preview.SetStatus(
                        $"You can't run as {UserName}: that needs IMPERSONATE permission on them (e.g. db_owner). " +
                        "You can run as yourself instead.",
                        Color.Firebrick);
                    runAsMe.Visible = true;
                }
            };

            runPage.Controls.Add(preview);
            UpdateRunHeading();
        }

        /// <summary>
        /// Whether the run can impersonate the end user. Not for Entra users:
        /// EXECUTE AS fails for them, so the run is as you.
        /// </summary>
        private bool CanRunAsUser => SelectedSignIn != null && SelectedSignIn != AccessSignIn.Entra;

        private string Me => string.IsNullOrEmpty(me) ? "your login" : me!;

        private void UpdateRunHeading()
        {
            string what = EffectiveProcedure ?? "the query";
            string runAs = runAsMe.Checked || !CanRunAsUser ? $"you ({Me})" : UserName ?? "";

            runHeading.Text = $"Run {what} as {runAs}";

            if (preview != null)
                preview.RunButton.Text = "Run as " + runAs;
        }

        private void Finish()
        {
            if (preview == null || !preview.HasResults)
                return;

            string? effective = EffectiveProcedure;

            if (!PreviewPanel.ConfirmFirstResultSetOnly(this, effective ?? "The query", preview.Results, "Excel (Power Query)"))
                return;

            var result = preview.Results[0];

            if (result.Truncated)
            {
                MessageBox.Show(
                    this,
                    $"The result has more than {PreviewPanel.MaxRows:N0} rows, too many to include.\n\n" +
                    "The workbook will start with just the column headings; the first Data > Refresh All loads all the data.",
                    "Flank",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            bool fromProcedure = effective != null;
            var parameters = effective == procedure && procedure != null ? procedureParameters : new List<SprocParameter>();

            Request = new WorkbookRequest(
                fromProcedure ? ProcedureCall.BuildExec(effective!, parameters) : query!,
                server,
                database)
            {
                ParameterNames = parameters.Where(p => !p.IsOutput).Select(p => p.Name.TrimStart('@')).ToList(),
                ParameterValues = preview.ParameterValues,
                ColumnNames = result.Columns,
                Rows = result.Truncated ? null : result.Rows,
                RowsFetchedAt = result.Truncated ? (DateTime?)null : preview.RanAt,
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        // ---------------------------------------------------------------- names

        private static string StripBrackets(string name) => name.Replace("[", "").Replace("]", "");

        /// <summary>"monthly", "dbo.monthly" or "[dbo].[monthly]" -> "[dbo].[monthly]"; null if invalid.</summary>
        private static string? NormalizeProcedureName(string text)
        {
            var parts = StripBrackets(text.Trim()).Split('.');

            if (parts.Length > 2 || parts.Any(p => string.IsNullOrWhiteSpace(p)))
                return null;

            string schema = parts.Length == 2 ? parts[0].Trim() : "dbo";
            string name = parts[parts.Length - 1].Trim();

            return AccessScript.QuoteName(schema) + "." + AccessScript.QuoteName(name);
        }
    }
}
