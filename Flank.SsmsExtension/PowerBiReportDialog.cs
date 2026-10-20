using Flank.PowerBi;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    internal sealed class PowerBiReportOptions
    {
        public PowerBiReportOptions(PowerBiWorkspace workspace, string reportName, string? sql, PowerBiConnection connection)
        {
            Workspace = workspace;
            ReportName = reportName;
            Sql = sql;
            Connection = connection;
        }

        public PowerBiWorkspace Workspace { get; }

        public string ReportName { get; }

        /// <summary>The query, edited by the user. Null for a stored procedure.</summary>
        public string? Sql { get; }

        /// <summary>The Gateway / Cloud Connection the report reads the database through.</summary>
        public PowerBiConnection Connection { get; }
    }

    /// <summary>
    /// Where to publish the paginated report, its name, and the Gateway /
    /// Cloud Connection it reads the database through. As with SSRS shared
    /// data sources, connections are set up in Power BI beforehand.
    /// </summary>
    internal sealed class PowerBiReportDialog : Form
    {
        private const string ManageConnectionsUrl = "https://app.powerbi.com/groups/me/gateways";

        private readonly string server;
        private readonly string database;
        private readonly string? preferredConnectionId;
        private readonly Func<Task<IReadOnlyList<PowerBiConnection>?>> allowConnections;
        private readonly ComboBox workspaceComboBox;
        private readonly TextBox reportNameTextBox;
        private readonly TextBox? sqlTextBox;
        private readonly ComboBox connectionComboBox;
        private readonly Label connectionHint;
        private readonly LinkLabel connectionLink;
        private readonly Button publishButton;
        private bool connectionsAllowed;

        public PowerBiReportOptions? Options { get; private set; }

        /// <param name="server">The server SSMS is connected to, to pick a connection to it by default.</param>
        /// <param name="connections">The Gateway / Cloud Connections, or null if Flank may not list them.</param>
        /// <param name="allowConnections">Signs in for listing connections; returns them, or null if that failed.</param>
        public PowerBiReportDialog(
            string signedInAs,
            string server,
            string database,
            IReadOnlyList<PowerBiWorkspace> workspaces,
            IReadOnlyList<PowerBiConnection>? connections,
            Func<Task<IReadOnlyList<PowerBiConnection>?>> allowConnections,
            string? sql,
            string? storedProcedure,
            string defaultReportName,
            PowerBiSettings settings)
        {
            this.server = server;
            this.database = database;
            this.allowConnections = allowConnections;
            preferredConnectionId = settings.ConnectionId;

            Text = "Create Power BI Paginated Report (Beta)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;

            Controls.Add(new Label
            {
                Text = $"Signed in to Power BI as {signedInAs}.",
                ForeColor = SystemColors.GrayText,
                Location = new Point(20, 16),
                Width = 600
            });

            Controls.Add(new Label { Text = "Workspace:", Location = new Point(20, 48), Width = 160 });

            workspaceComboBox = new ComboBox
            {
                Location = new Point(185, 44),
                Width = 435,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            workspaceComboBox.Items.AddRange(workspaces.Cast<object>().ToArray());
            workspaceComboBox.SelectedItem = workspaces.FirstOrDefault(w => w.Id == settings.WorkspaceId);
            if (workspaceComboBox.SelectedIndex < 0 && workspaceComboBox.Items.Count > 0)
                workspaceComboBox.SelectedIndex = 0;

            Controls.Add(workspaceComboBox);

            reportNameTextBox = AddField("Report Name:", defaultReportName, 82);

            Controls.Add(new Label { Text = "Gateway / Cloud Connection:", Location = new Point(20, 120), Width = 165 });

            connectionComboBox = new ComboBox
            {
                Location = new Point(185, 116),
                Width = 435,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            connectionHint = new Label
            {
                ForeColor = SystemColors.GrayText,
                Location = new Point(185, 142),
                Size = new Size(435, 32)
            };

            connectionLink = new LinkLabel
            {
                Location = new Point(185, 174),
                Width = 435
            };
            connectionLink.LinkClicked += async (s, e) => await ConnectionLinkClickedAsync();

            Controls.Add(connectionComboBox);
            Controls.Add(connectionHint);
            Controls.Add(connectionLink);

            int y;

            if (storedProcedure != null)
            {
                AddField("Stored Procedure:", storedProcedure, 204).ReadOnly = true;
                y = 248;
            }
            else
            {
                Controls.Add(new Label { Text = "SQL:", Location = new Point(20, 204), Width = 160 });

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
                    Location = new Point(20, 224),
                    Size = new Size(600, 140)
                };

                Controls.Add(sqlTextBox);
                y = 378;
            }

            ClientSize = new Size(640, y + 40);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(450, y),
                Width = 80
            };

            publishButton = new Button
            {
                Text = "Publish",
                Location = new Point(540, y),
                Width = 80
            };

            publishButton.Click += PublishButton_Click;

            Controls.Add(cancelButton);
            Controls.Add(publishButton);

            AcceptButton = publishButton;
            CancelButton = cancelButton;

            ShowConnections(connections);
        }

        private void ShowConnections(IReadOnlyList<PowerBiConnection>? connections)
        {
            connectionsAllowed = connections != null;
            connectionComboBox.Items.Clear();

            if (connections == null)
            {
                connectionComboBox.Enabled = false;
                connectionHint.Text = "Flank needs permission to list your Power BI connections.";
                connectionLink.Text = "Allow Flank to list connections (sign in again)";
                publishButton.Enabled = false;
                return;
            }

            connectionComboBox.Items.AddRange(connections.Cast<object>().ToArray());

            if (connections.Count == 0)
            {
                connectionComboBox.Enabled = false;
                connectionHint.Text =
                    "You have no SQL Server connections in Power BI. Add one for this database, then try again.";
                connectionLink.Text = "Open Manage connections and gateways";
                publishButton.Enabled = false;
                return;
            }

            // The last one used if it's for this database; else the first for it; else the last one used.
            var forThisDatabase = connections.Where(c => c.IsFor(server, database)).ToList();
            var previous = connections.FirstOrDefault(c => c.Id == preferredConnectionId);

            connectionComboBox.SelectedItem =
                (previous != null && previous.IsFor(server, database) ? previous : null) ??
                forThisDatabase.FirstOrDefault() ??
                previous ??
                connections[0];

            connectionComboBox.Enabled = true;
            publishButton.Enabled = true;
            connectionHint.Text = forThisDatabase.Count > 0
                ? "Connections are set up in Power BI, under Manage connections and gateways."
                : $"None of your connections is for {database} on {server}. Add one, or pick another.";
            connectionLink.Text = "Open Manage connections and gateways";
        }

        private async Task ConnectionLinkClickedAsync()
        {
            if (connectionsAllowed)
            {
                try
                {
                    Process.Start(ManageConnectionsUrl);
                }
                catch (Exception)
                {
                    MessageBox.Show($"Couldn't open a browser. Go to {ManageConnectionsUrl}", "Flank");
                }

                return;
            }

            connectionLink.Enabled = false;
            var allowed = await allowConnections();
            connectionLink.Enabled = true;

            if (allowed != null)
                ShowConnections(allowed);
        }

        private TextBox AddField(string label, string value, int y)
        {
            Controls.Add(new Label { Text = label, Location = new Point(20, y + 4), Width = 160 });

            var textBox = new TextBox { Text = value, Location = new Point(185, y), Width = 435 };
            Controls.Add(textBox);
            return textBox;
        }

        private void PublishButton_Click(object sender, EventArgs e)
        {
            string? problem = null;

            if (!(workspaceComboBox.SelectedItem is PowerBiWorkspace))
                problem = "Choose a workspace. If the list is empty, create one in Power BI first.";
            else if (string.IsNullOrWhiteSpace(reportNameTextBox.Text))
                problem = "Enter a report name.";
            else if (sqlTextBox != null && string.IsNullOrWhiteSpace(sqlTextBox.Text))
                problem = "The SQL is empty.";
            else if (!(connectionComboBox.SelectedItem is PowerBiConnection))
                problem = "Choose a Gateway / Cloud Connection.";

            if (problem != null)
            {
                MessageBox.Show(problem, "Flank", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Options = new PowerBiReportOptions(
                (PowerBiWorkspace)workspaceComboBox.SelectedItem!,
                reportNameTextBox.Text.Trim(),
                sqlTextBox?.Text,
                (PowerBiConnection)connectionComboBox.SelectedItem!);

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
