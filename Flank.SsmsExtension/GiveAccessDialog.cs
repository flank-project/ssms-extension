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
    /// Builds (and optionally runs) the T-SQL that lets an end user refresh a
    /// Flank workbook or report: create their user, grant the narrowest access.
    /// </summary>
    internal sealed class GiveAccessDialog : Form
    {
        private readonly string server;
        private readonly string database;
        private readonly bool isAzureSqlDatabase;
        private readonly string? procedure;
        private readonly IReadOnlyList<string> queryTables;
        private readonly Action<Action<IDbConnection>> useConnection;

        private readonly ComboBox signIn;
        private readonly Label nameLabel;
        private readonly TextBox name;
        private readonly Label passwordLabel;
        private readonly TextBox password;
        private readonly Button generate;
        private readonly RadioButton procedureScope;
        private readonly RadioButton objectsScope;
        private readonly RadioButton allTablesScope;
        private readonly TextBox script;
        private readonly Label instructions;
        private readonly Label status;
        private readonly Button runButton;
        private readonly Button testButton;

        private readonly List<AccessSignIn> signInOptions;

        /// <param name="procedure">[schema].[name] when giving access for a stored procedure.</param>
        /// <param name="queryTables">Tables a query reads, when giving access for a query.</param>
        /// <param name="useConnection">
        /// Calls the given action with a connection to the database. The
        /// connection may be borrowed (e.g. the query window's), so the dialog
        /// never closes it.
        /// </param>
        public GiveAccessDialog(
            string server,
            string database,
            bool isAzureSqlDatabase,
            string? procedure,
            IReadOnlyList<string> queryTables,
            Action<Action<IDbConnection>> useConnection)
        {
            this.server = server;
            this.database = database;
            this.isAzureSqlDatabase = isAzureSqlDatabase;
            this.procedure = procedure;
            this.queryTables = queryTables;
            this.useConnection = useConnection;

            Text = "Give someone access - " + database;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(720, 640);
            MinimumSize = new Size(620, 560);

            int y = 16;

            Controls.Add(new Label
            {
                Text = "Create a database user for the person who'll refresh the workbook or report, " +
                       "and give them only the access it needs.",
                Location = new Point(16, y),
                Size = new Size(688, 34),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            });
            y += 44;

            // Azure SQL Database has no Windows logins; SQL Server (on-premises)
            // users don't normally have Entra accounts in the database.
            signInOptions = isAzureSqlDatabase
                ? new List<AccessSignIn> { AccessSignIn.SqlLogin, AccessSignIn.Entra }
                : new List<AccessSignIn> { AccessSignIn.SqlLogin, AccessSignIn.Windows };

            AddLabel("Signs in with:", y);
            signIn = new ComboBox
            {
                Location = new Point(150, y - 3),
                Width = 400,
                DropDownStyle = ComboBoxStyle.DropDownList,
            };
            foreach (var option in signInOptions)
                signIn.Items.Add(Describe(option));
            signIn.SelectedIndex = 0;
            signIn.SelectedIndexChanged += (s, e) => UpdateForm();
            Controls.Add(signIn);
            y += 34;

            nameLabel = AddLabel("Username:", y);
            name = new TextBox { Location = new Point(150, y - 3), Width = 400, Name = "AccessName" };
            name.TextChanged += (s, e) => UpdateForm();
            Controls.Add(name);
            y += 34;

            passwordLabel = AddLabel("Password:", y);
            password = new TextBox { Location = new Point(150, y - 3), Width = 300, Text = AccessScript.GeneratePassword(), Name = "AccessPassword" };
            password.TextChanged += (s, e) => UpdateForm();
            Controls.Add(password);
            generate = new Button { Text = "New password", Location = new Point(460, y - 4), Width = 90 };
            generate.Click += (s, e) => password.Text = AccessScript.GeneratePassword();
            Controls.Add(generate);
            y += 40;

            AddLabel("Access:", y);
            procedureScope = AddScope($"Only this stored procedure ({procedure})", y, procedure != null);
            if (procedure != null) y += 26;
            objectsScope = AddScope(
                "Only the tables this query reads: " + string.Join(", ", queryTables),
                y,
                procedure == null && queryTables.Count > 0);
            if (procedure == null && queryTables.Count > 0) y += 26;
            allTablesScope = AddScope("Every table and view in the database (db_datareader)", y, true);
            y += 34;

            (procedure != null ? procedureScope : queryTables.Count > 0 ? objectsScope : allTablesScope).Checked = true;

            AddLabel("Script:", y);
            y += 22;
            script = new TextBox
            {
                Location = new Point(16, y),
                Size = new Size(688, 170),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Name = "AccessScript",
            };
            Controls.Add(script);

            instructions = new Label
            {
                Location = new Point(16, ClientSize.Height - 150),
                Size = new Size(688, 80),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(instructions);

            status = new Label
            {
                Location = new Point(16, ClientSize.Height - 66),
                Size = new Size(688, 20),
                AutoEllipsis = true,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(status);

            var copyButton = new Button { Text = "Copy script", Width = 100, Location = new Point(ClientSize.Width - 466, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            // Checks the user's permissions by impersonating them.
            testButton = new Button { Text = "Test as this user", Width = 120, Location = new Point(ClientSize.Width - 346, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            testButton.Click += (s, e) => Test();
            copyButton.Click += (s, e) =>
            {
                if (TryBuild(out var request))
                    Clipboard.SetText(AccessScript.ToText(request));
                status.ForeColor = SystemColors.ControlText;
                status.Text = "Script copied.";
            };

            runButton = new Button { Text = "Run script", Width = 100, Location = new Point(ClientSize.Width - 226, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            runButton.Click += (s, e) => Run();

            var closeButton = new Button { Text = "Close", Width = 100, DialogResult = DialogResult.Cancel, Location = new Point(ClientSize.Width - 116, ClientSize.Height - 40), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            Controls.Add(copyButton);
            Controls.Add(testButton);
            Controls.Add(runButton);
            Controls.Add(closeButton);
            CancelButton = closeButton;

            UpdateForm();
        }

        private AccessSignIn SelectedSignIn => signInOptions[signIn.SelectedIndex];

        private static string Describe(AccessSignIn signIn)
        {
            switch (signIn)
            {
                case AccessSignIn.Entra: return "Microsoft Entra account";
                case AccessSignIn.Windows: return "Windows / Active Directory account";
                default: return "SQL Server username and password";
            }
        }

        private Label AddLabel(string text, int y)
        {
            var label = new Label { Text = text, Location = new Point(16, y), Width = 130 };
            Controls.Add(label);
            return label;
        }

        private RadioButton AddScope(string text, int y, bool visible)
        {
            var radio = new RadioButton
            {
                Text = text,
                Location = new Point(150, y - 3),
                Size = new Size(554, 24),
                AutoEllipsis = true,
                Visible = visible,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            radio.CheckedChanged += (s, e) => UpdateForm();
            Controls.Add(radio);
            return radio;
        }

        private bool TryBuild(out AccessRequest request)
        {
            request = new AccessRequest
            {
                SignIn = SelectedSignIn,
                Name = name.Text.Trim(),
                Password = password.Text,
                Database = database,
                IsAzureSqlDatabase = isAzureSqlDatabase,
                Scope = procedureScope.Checked ? AccessScope.Procedure
                      : objectsScope.Checked ? AccessScope.Objects
                      : AccessScope.AllTables,
                Procedure = procedure ?? "",
                Objects = queryTables,
            };

            return request.Name.Length > 0 &&
                   (request.SignIn != AccessSignIn.SqlLogin || request.Password.Length > 0);
        }

        private void UpdateForm()
        {
            // Change events fire while the form is still being built.
            if (script == null || instructions == null)
                return;

            var selected = SelectedSignIn;
            bool sqlLogin = selected == AccessSignIn.SqlLogin;

            nameLabel.Text = selected == AccessSignIn.Entra ? "Email (UPN):"
                           : selected == AccessSignIn.Windows ? @"DOMAIN\user:"
                           : "Username:";
            passwordLabel.Visible = password.Visible = generate.Visible = sqlLogin;

            bool ready = TryBuild(out var request);
            script.Text = ready ? AccessScript.ToText(request) : "-- Enter a name to see the script.";
            runButton.Enabled = ready;
            testButton.Enabled = ready;

            string who = ready ? request.Name : "the user";
            switch (selected)
            {
                case AccessSignIn.SqlLogin:
                    instructions.Text =
                        $"Send them the workbook, plus: server {server}, database {database}, username {who} and the password above.\n" +
                        "When Excel asks them to sign in, they choose Database and enter that username and password.";
                    break;
                case AccessSignIn.Entra:
                    instructions.Text =
                        $"{who} must already be in your Microsoft Entra tenant, and the server needs an Entra admin. " +
                        "Run this from a connection signed in with Entra (SQL logins can't create Entra users).\n" +
                        "When Excel asks them to sign in, they choose Microsoft account and use that Entra account.";
                    break;
                case AccessSignIn.Windows:
                    instructions.Text =
                        $"Their computer must be able to reach {server} (e.g. on the network or VPN).\n" +
                        "When Excel asks them to sign in, they choose Windows and use their current credentials.";
                    break;
            }
        }

        /// <summary>
        /// Whether the user can do what the workbook or report needs: EXECUTE on
        /// the procedure, or SELECT on each table the query reads.
        /// </summary>
        private void Test()
        {
            string user = name.Text.Trim();

            var checks = procedure != null
                ? new List<PermissionCheck> { new PermissionCheck(procedure, "EXECUTE") }
                : queryTables.Select(t => new PermissionCheck(t, "SELECT")).ToList();

            if (checks.Count == 0)
            {
                status.ForeColor = SystemColors.ControlText;
                status.Text = "Couldn't tell which tables this query reads, so there's nothing to test.";
                return;
            }

            status.ForeColor = SystemColors.ControlText;
            status.Text = "Testing...";
            status.Refresh();
            UseWaitCursor = true;

            try
            {
                bool exists = false;

                useConnection(connection =>
                {
                    exists = AccessTester.UserExists(connection, user);
                    if (exists)
                        AccessTester.Check(connection, user, checks);
                });

                if (!exists)
                {
                    status.ForeColor = Color.Firebrick;
                    status.Text = $"{user} doesn't exist yet. Run the script first.";
                    return;
                }

                var missing = checks.Where(c => !c.Granted).ToList();

                if (missing.Count == 0)
                {
                    status.ForeColor = Color.ForestGreen;
                    status.Text = $"{user} can " + string.Join(", ", checks.Select(c => $"{c.Permission} {c.ObjectName}")) + ".";
                }
                else
                {
                    status.ForeColor = Color.Firebrick;
                    status.Text = $"{user} can't " + string.Join(", ", missing.Select(c => $"{c.Permission} {c.ObjectName}")) +
                                  ". Run the script, or choose broader access.";
                }
            }
            catch (Exception ex)
            {
                status.ForeColor = Color.Firebrick;
                // The user exists (checked above), so a failure to impersonate
                // means we lack permission to.
                status.Text = ex.Message.IndexOf("Cannot execute as", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "You need IMPERSONATE permission on that user (e.g. db_owner) to test as them."
                    : ex.Message;
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        // Runs on the UI thread: the statements are quick, and the connection
        // may be the query window's own.
        private void Run()
        {
            if (!TryBuild(out var request))
                return;

            runButton.Enabled = false;
            status.ForeColor = SystemColors.ControlText;
            status.Text = "Running...";
            status.Refresh();
            UseWaitCursor = true;

            try
            {
                useConnection(connection => AccessScript.Run(connection, request));

                status.ForeColor = Color.ForestGreen;
                status.Text = $"Done: {request.Name} has access.";
            }
            catch (Exception ex)
            {
                status.ForeColor = Color.Firebrick;
                status.Text = ex.Message;
            }
            finally
            {
                UseWaitCursor = false;
                runButton.Enabled = true;
            }
        }
    }
}
