using Flank.PowerBi;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class PowerBiReportOptions
    {
        public PowerBiReportOptions(PowerBiWorkspace workspace, string reportName, string? sql, PowerBiDataAccess access)
        {
            Workspace = workspace;
            ReportName = reportName;
            Sql = sql;
            Access = access;
        }

        public PowerBiWorkspace Workspace { get; }

        public string ReportName { get; }

        /// <summary>The query, edited by the user. Null for a stored procedure.</summary>
        public string? Sql { get; }

        public PowerBiDataAccess Access { get; }
    }

    /// <summary>
    /// Where to publish the paginated report, its name, and how the people
    /// who run it read the database: a SQL login saved on the report, or a
    /// shared connection (Power BI's equivalent of an SSRS shared data source).
    /// </summary>
    internal sealed class PowerBiReportDialog : Form
    {
        private sealed class NewConnectionItem
        {
            public override string ToString() => "New shared connection (with the login below)";
        }

        private readonly string server;
        private readonly string database;
        private readonly Func<Task<IReadOnlyList<PowerBiConnection>?>> allowConnections;
        private readonly ComboBox workspaceComboBox;
        private readonly TextBox reportNameTextBox;
        private readonly TextBox? sqlTextBox;
        private readonly RadioButton reportLoginRadio;
        private readonly RadioButton connectionRadio;
        private readonly ComboBox connectionComboBox;
        private readonly LinkLabel allowConnectionsLink;
        private readonly TextBox usernameTextBox;
        private readonly TextBox passwordTextBox;

        public PowerBiReportOptions? Options { get; private set; }

        /// <param name="connections">Shared connections to this database, or null if Flank may not list them.</param>
        /// <param name="allowConnections">Signs in for shared connections; returns them, or null if that failed.</param>
        /// <param name="login">The SQL login of the SSMS connection, if it uses one.</param>
        public PowerBiReportDialog(
            string signedInAs,
            string server,
            string database,
            IReadOnlyList<PowerBiWorkspace> workspaces,
            IReadOnlyList<PowerBiConnection>? connections,
            Func<Task<IReadOnlyList<PowerBiConnection>?>> allowConnections,
            SsmsSqlLogin? login,
            string? sql,
            string? storedProcedure,
            string defaultReportName,
            PowerBiSettings settings)
        {
            this.server = server;
            this.database = database;
            this.allowConnections = allowConnections;

            Text = "Create Power BI Paginated Report";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;

            Controls.Add(new Label
            {
                Text = $"Signed in to Power BI as {signedInAs}.",
                Location = new Point(20, 20),
                Width = 500
            });

            Controls.Add(new Label { Text = "Workspace:", Location = new Point(20, 54), Width = 130 });

            workspaceComboBox = new ComboBox
            {
                Location = new Point(155, 50),
                Width = 365,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            workspaceComboBox.Items.AddRange(workspaces.Cast<object>().ToArray());
            workspaceComboBox.SelectedItem = workspaces.FirstOrDefault(w => w.Id == settings.WorkspaceId);
            if (workspaceComboBox.SelectedIndex < 0 && workspaceComboBox.Items.Count > 0)
                workspaceComboBox.SelectedIndex = 0;

            Controls.Add(workspaceComboBox);

            reportNameTextBox = AddField("Report Name:", defaultReportName, 90);

            int y;

            if (storedProcedure != null)
            {
                AddField("Stored Procedure:", storedProcedure, 130).ReadOnly = true;
                y = 175;
            }
            else
            {
                Controls.Add(new Label { Text = "SQL:", Location = new Point(20, 130), Width = 130 });

                sqlTextBox = new TextBox
                {
                    Text = sql,
                    Multiline = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    AcceptsReturn = true,
                    AcceptsTab = true,
                    Font = new Font("Consolas", 9),
                    Location = new Point(20, 150),
                    Size = new Size(500, 130)
                };

                Controls.Add(sqlTextBox);
                y = 295;
            }

            Controls.Add(new Label
            {
                Text = "When people run the report, it reads the database using:",
                Location = new Point(20, y),
                Width = 500
            });

            reportLoginRadio = new RadioButton
            {
                Text = "A SQL login saved on this report",
                Location = new Point(30, y + 25),
                Width = 400,
                Checked = true
            };

            connectionRadio = new RadioButton
            {
                Text = "A shared connection:",
                Location = new Point(30, y + 52),
                Width = 150
            };

            connectionComboBox = new ComboBox
            {
                Location = new Point(185, y + 51),
                Width = 335,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            allowConnectionsLink = new LinkLabel
            {
                Text = "Let Flank list and create shared connections (sign in again)",
                Location = new Point(48, y + 80),
                Width = 470
            };

            allowConnectionsLink.LinkClicked += async (sender, e) =>
            {
                allowConnectionsLink.Enabled = false;
                var allowed = await this.allowConnections();
                allowConnectionsLink.Enabled = true;

                if (allowed != null)
                    ShowConnections(allowed, null);
            };

            Controls.Add(new Label { Text = "SQL login:", Location = new Point(30, y + 114), Width = 120 });

            usernameTextBox = new TextBox
            {
                Text = login?.Username ?? "",
                Location = new Point(155, y + 110),
                Width = 365
            };

            Controls.Add(new Label { Text = "Password:", Location = new Point(30, y + 144), Width = 120 });

            passwordTextBox = new TextBox
            {
                Text = login?.Password ?? "",
                UseSystemPasswordChar = true,
                Location = new Point(155, y + 140),
                Width = 365
            };

            var loginHint = new Label
            {
                Text = login != null
                    ? "From your SSMS connection. Power BI keeps it encrypted; it's never put in the report."
                    : "Your SSMS connection doesn't use a SQL login, so enter one for the report to use.",
                Location = new Point(155, y + 166),
                Size = new Size(365, 32),
                ForeColor = SystemColors.GrayText
            };

            Controls.Add(reportLoginRadio);
            Controls.Add(connectionRadio);
            Controls.Add(connectionComboBox);
            Controls.Add(allowConnectionsLink);
            Controls.Add(usernameTextBox);
            Controls.Add(passwordTextBox);
            Controls.Add(loginHint);

            reportLoginRadio.CheckedChanged += (sender, e) => UpdateEnabled();
            connectionComboBox.SelectedIndexChanged += (sender, e) => UpdateEnabled();

            int buttonsY = y + 210;
            ClientSize = new Size(540, buttonsY + 40);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(350, buttonsY),
                Width = 80
            };

            var publishButton = new Button
            {
                Text = "Publish",
                Location = new Point(440, buttonsY),
                Width = 80
            };

            publishButton.Click += PublishButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(publishButton);

            AcceptButton = publishButton;
            CancelButton = cancelButton;

            ShowConnections(connections, settings.ConnectionId);

            if (connections != null && settings.UseSharedConnection)
                connectionRadio.Checked = true;

            UpdateEnabled();
        }

        private void ShowConnections(IReadOnlyList<PowerBiConnection>? connections, string? preferredId)
        {
            connectionComboBox.Items.Clear();

            if (connections == null)
            {
                connectionRadio.Enabled = false;
                connectionComboBox.Enabled = false;
                allowConnectionsLink.Visible = true;
                return;
            }

            connectionComboBox.Items.AddRange(connections.Cast<object>().ToArray());
            connectionComboBox.Items.Add(new NewConnectionItem());
            connectionComboBox.SelectedItem =
                (object?)connections.FirstOrDefault(c => c.Id == preferredId) ??
                (object?)connections.FirstOrDefault() ??
                connectionComboBox.Items[connectionComboBox.Items.Count - 1];

            connectionRadio.Enabled = true;
            allowConnectionsLink.Visible = false;
            UpdateEnabled();
        }

        // The login is used for a login on the report, or to create a connection.
        private void UpdateEnabled()
        {
            bool needsLogin = reportLoginRadio.Checked || connectionComboBox.SelectedItem is NewConnectionItem;
            connectionComboBox.Enabled = connectionRadio.Enabled && connectionRadio.Checked;
            usernameTextBox.Enabled = needsLogin;
            passwordTextBox.Enabled = needsLogin;
        }

        private TextBox AddField(string label, string value, int y)
        {
            Controls.Add(new Label { Text = label, Location = new Point(20, y + 4), Width = 130 });

            var textBox = new TextBox { Text = value, Location = new Point(155, y), Width = 365 };
            Controls.Add(textBox);
            return textBox;
        }

        private void PublishButton_Click(object sender, EventArgs e)
        {
            string? problem = null;
            bool needsLogin = reportLoginRadio.Checked || connectionComboBox.SelectedItem is NewConnectionItem;

            if (!(workspaceComboBox.SelectedItem is PowerBiWorkspace workspace))
                problem = "Choose a workspace. If the list is empty, create one in Power BI first.";
            else if (string.IsNullOrWhiteSpace(reportNameTextBox.Text))
                problem = "Enter a report name.";
            else if (sqlTextBox != null && string.IsNullOrWhiteSpace(sqlTextBox.Text))
                problem = "The SQL is empty.";
            else if (connectionRadio.Checked && connectionComboBox.SelectedItem == null)
                problem = "Choose a shared connection.";
            else if (needsLogin && (string.IsNullOrWhiteSpace(usernameTextBox.Text) || passwordTextBox.Text.Length == 0))
                problem = "Enter the SQL login and password the report should use.";

            if (problem != null)
            {
                MessageBox.Show(problem, "Flank", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string username = usernameTextBox.Text.Trim();
            string password = passwordTextBox.Text;

            PowerBiDataAccess access =
                reportLoginRadio.Checked
                    ? new PowerBiDataAccess.ReportLogin(username, password)
                    : connectionComboBox.SelectedItem is PowerBiConnection existing
                        ? new PowerBiDataAccess.ExistingConnection(existing.Id)
                        : (PowerBiDataAccess)new PowerBiDataAccess.NewConnection(
                            $"{database} on {server} ({username})", username, password);

            Options = new PowerBiReportOptions(
                (PowerBiWorkspace)workspaceComboBox.SelectedItem!,
                reportNameTextBox.Text.Trim(),
                sqlTextBox?.Text,
                access);

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
