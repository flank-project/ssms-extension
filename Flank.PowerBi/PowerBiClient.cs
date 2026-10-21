using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Flank.PowerBi
{
    public class PowerBiException : Exception
    {
        public PowerBiException(string message, Exception? innerException = null)
            : base(message, innerException)
        {
        }
    }

    public sealed class PowerBiWorkspace
    {
        public PowerBiWorkspace(string id, string name)
        {
            Id = id;
            Name = name;
        }

        public string Id { get; }

        public string Name { get; }

        public override string ToString() => Name;
    }

    public sealed class PowerBiReport
    {
        public PowerBiReport(string id, string name)
        {
            Id = id;
            Name = name;
        }

        public string Id { get; }

        public string Name { get; }
    }

    /// <summary>A data source of a paginated report, as the service resolved it.</summary>
    public sealed class PowerBiReportDataSource
    {
        public PowerBiReportDataSource(string name, string gatewayId, string datasourceId, string server, string database)
        {
            Name = name;
            GatewayId = gatewayId;
            DatasourceId = datasourceId;
            Server = server;
            Database = database;
        }

        public string Name { get; }

        public string GatewayId { get; }

        public string DatasourceId { get; }

        public string Server { get; }

        public string Database { get; }
    }

    /// <summary>A shareable cloud connection: Power BI's equivalent of an SSRS shared data source.</summary>
    public sealed class PowerBiConnection
    {
        public PowerBiConnection(string id, string displayName)
        {
            Id = id;
            DisplayName = displayName;
        }

        public string Id { get; }

        public string DisplayName { get; }

        public override string ToString() => DisplayName;
    }

    /// <summary>
    /// The Power BI and Fabric REST calls Flank needs to publish a paginated
    /// report. Tokens come from the caller: the user's (PowerBiSignIn) in
    /// SSMS, a service principal's in tests.
    /// </summary>
    public sealed class PowerBiClient
    {
        private const string PowerBiApi = "https://api.powerbi.com/v1.0/myorg";
        private const string FabricApi = "https://api.fabric.microsoft.com/v1";

        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

        private readonly Func<CancellationToken, Task<string>> getPowerBiToken;
        private readonly Func<CancellationToken, Task<string?>>? getConnectionsToken;

        /// <param name="getPowerBiToken">An access token for the Power BI API.</param>
        /// <param name="getConnectionsToken">
        /// An access token for the Fabric connections API, or null when the
        /// caller hasn't been granted it (shared connections are then unavailable).
        /// </param>
        public PowerBiClient(
            Func<CancellationToken, Task<string>> getPowerBiToken,
            Func<CancellationToken, Task<string?>>? getConnectionsToken = null)
        {
            this.getPowerBiToken = getPowerBiToken;
            this.getConnectionsToken = getConnectionsToken;
        }

        public static string GetReportUrl(string workspaceId, string reportId) =>
            $"https://app.powerbi.com/groups/{workspaceId}/rdlreports/{reportId}";

        public async Task<IReadOnlyList<PowerBiWorkspace>> GetWorkspacesAsync(CancellationToken cancellationToken = default)
        {
            var json = await SendPowerBiAsync(HttpMethod.Get, "/groups", null, cancellationToken).ConfigureAwait(false);

            return json["value"]!
                .Select(g => new PowerBiWorkspace((string)g["id"]!, (string)g["name"]!))
                .OrderBy(w => w.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>The workspace's paginated reports.</summary>
        public async Task<IReadOnlyList<PowerBiReport>> GetReportsAsync(string workspaceId, CancellationToken cancellationToken = default)
        {
            var json = await SendPowerBiAsync(HttpMethod.Get, $"/groups/{workspaceId}/reports", null, cancellationToken).ConfigureAwait(false);

            return json["value"]!
                .Where(r => (string?)r["reportType"] == "PaginatedReport")
                .Select(r => new PowerBiReport((string)r["id"]!, (string?)r["name"] ?? ""))
                .ToList();
        }

        /// <summary>The report's id, or null if the workspace has no paginated report by that name.</summary>
        public async Task<string?> FindReportAsync(string workspaceId, string reportName, CancellationToken cancellationToken = default)
        {
            return (await GetReportsAsync(workspaceId, cancellationToken).ConfigureAwait(false))
                .FirstOrDefault(r => string.Equals(r.Name, reportName, StringComparison.OrdinalIgnoreCase))
                ?.Id;
        }

        /// <summary>
        /// Uploads a paginated report, replacing one with the same name. Returns its id.
        /// </summary>
        public async Task<string> ImportReportAsync(
            string workspaceId,
            string reportName,
            string rdl,
            CancellationToken cancellationToken = default)
        {
            // Overwrite fails when there's nothing to overwrite, and Abort when there is.
            string? existing = await FindReportAsync(workspaceId, reportName, cancellationToken).ConfigureAwait(false);
            string conflict = existing != null ? "Overwrite" : "Abort";

            var content = new MultipartFormDataContent();
            var file = new ByteArrayContent(Encoding.UTF8.GetBytes(rdl));
            file.Headers.ContentType = new MediaTypeHeaderValue("application/xml");
            content.Add(file, "file", reportName + ".rdl");

            string path = $"/groups/{workspaceId}/imports" +
                $"?datasetDisplayName={Uri.EscapeDataString(reportName + ".rdl")}&nameConflict={conflict}";

            var import = await SendPowerBiAsync(HttpMethod.Post, path, content, cancellationToken).ConfigureAwait(false);
            string importId = (string)import["id"]!;

            for (int attempt = 0; attempt < 60; attempt++)
            {
                await Task.Delay(TimeSpan.FromSeconds(attempt == 0 ? 1 : 2), cancellationToken).ConfigureAwait(false);

                var status = await SendPowerBiAsync(HttpMethod.Get, $"/groups/{workspaceId}/imports/{importId}", null, cancellationToken).ConfigureAwait(false);
                string state = (string)status["importState"]!;

                if (state == "Succeeded")
                    return (string)status["reports"]![0]!["id"]!;

                if (state == "Failed")
                    throw new PowerBiException("Power BI couldn't import the report: " + DescribeError(status["error"]));
            }

            throw new PowerBiException("Power BI is still importing the report after two minutes.");
        }

        public async Task<IReadOnlyList<PowerBiReportDataSource>> GetReportDataSourcesAsync(
            string workspaceId,
            string reportId,
            CancellationToken cancellationToken = default)
        {
            var json = await SendPowerBiAsync(HttpMethod.Get, $"/groups/{workspaceId}/reports/{reportId}/datasources", null, cancellationToken).ConfigureAwait(false);

            return json["value"]!
                .Select(d => new PowerBiReportDataSource(
                    (string?)d["name"] ?? "",
                    (string?)d["gatewayId"] ?? "",
                    (string?)d["datasourceId"] ?? "",
                    (string?)d["connectionDetails"]?["server"] ?? "",
                    (string?)d["connectionDetails"]?["database"] ?? ""))
                .ToList();
        }

        /// <summary>
        /// Saves a SQL login on the report's own data source (a personal cloud
        /// connection). Everyone who runs the report reads the database as that login.
        /// </summary>
        public async Task SetSqlLoginAsync(
            PowerBiReportDataSource dataSource,
            string username,
            string password,
            CancellationToken cancellationToken = default)
        {
            string credentials = JsonConvert.SerializeObject(new
            {
                credentialData = new[]
                {
                    new { name = "username", value = username },
                    new { name = "password", value = password }
                }
            });

            var body = new
            {
                credentialDetails = new
                {
                    credentialType = "Basic",
                    credentials,
                    encryptedConnection = "Encrypted",
                    encryptionAlgorithm = "None",
                    privacyLevel = "None"
                }
            };

            await SendPowerBiAsync(
                new HttpMethod("PATCH"),
                $"/gateways/{dataSource.GatewayId}/datasources/{dataSource.DatasourceId}",
                Json(body),
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Whether the caller may list and create shared connections.</summary>
        public async Task<bool> CanUseConnectionsAsync(CancellationToken cancellationToken = default)
        {
            if (getConnectionsToken == null)
                return false;

            return await getConnectionsToken(cancellationToken).ConfigureAwait(false) != null;
        }

        /// <summary>The shared (shareable cloud) SQL connections to this server and database.</summary>
        public async Task<IReadOnlyList<PowerBiConnection>> GetSqlConnectionsAsync(
            string server,
            string database,
            CancellationToken cancellationToken = default)
        {
            var connections = new List<PowerBiConnection>();
            string? continuation = null;

            do
            {
                string path = "/connections" +
                    (continuation != null ? "?continuationToken=" + Uri.EscapeDataString(continuation) : "");

                var json = await SendFabricAsync(HttpMethod.Get, path, null, cancellationToken).ConfigureAwait(false);

                foreach (var c in json["value"]!)
                {
                    if ((string?)c["connectivityType"] != "ShareableCloud" ||
                        !string.Equals((string?)c["connectionDetails"]?["type"], "SQL", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // The path is "server;database".
                    string[] parts = ((string?)c["connectionDetails"]?["path"] ?? "").Split(';');

                    if (parts.Length == 2 &&
                        string.Equals(parts[0].Trim(), server, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(parts[1].Trim(), database, StringComparison.OrdinalIgnoreCase))
                    {
                        connections.Add(new PowerBiConnection((string)c["id"]!, (string?)c["displayName"] ?? ""));
                    }
                }

                continuation = (string?)json["continuationToken"];
            }
            while (!string.IsNullOrEmpty(continuation));

            return connections
                .OrderBy(c => c.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        /// <summary>Creates a shared SQL connection with a SQL login.</summary>
        public async Task<PowerBiConnection> CreateSqlConnectionAsync(
            string displayName,
            string server,
            string database,
            string username,
            string password,
            CancellationToken cancellationToken = default)
        {
            var body = new
            {
                connectivityType = "ShareableCloud",
                displayName,
                connectionDetails = new
                {
                    type = "SQL",
                    creationMethod = "Sql",
                    parameters = new[]
                    {
                        new { dataType = "Text", name = "server", value = server },
                        new { dataType = "Text", name = "database", value = database }
                    }
                },
                privacyLevel = "Organizational",
                credentialDetails = new
                {
                    singleSignOnType = "None",
                    connectionEncryption = "Encrypted",
                    skipTestConnection = false,
                    credentials = new { credentialType = "Basic", username, password }
                }
            };

            var json = await SendFabricAsync(HttpMethod.Post, "/connections", Json(body), cancellationToken).ConfigureAwait(false);
            return new PowerBiConnection((string)json["id"]!, (string?)json["displayName"] ?? displayName);
        }

        /// <summary>
        /// Points the report's data source at a shared connection. (Documented for
        /// on-premises gateways only, but it works for shareable cloud
        /// connections when given the report's own cloud gateway.)
        /// </summary>
        public async Task BindToConnectionAsync(
            string workspaceId,
            string reportId,
            PowerBiReportDataSource dataSource,
            string connectionId,
            CancellationToken cancellationToken = default)
        {
            var body = new
            {
                gatewayObjectId = dataSource.GatewayId,
                bindDetails = new[] { new { dataSourceName = dataSource.Name, dataSourceObjectId = connectionId } }
            };

            await SendPowerBiAsync(
                HttpMethod.Post,
                $"/groups/{workspaceId}/reports/{reportId}/Default.BindToGateway",
                Json(body),
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Renders the report to a file (e.g. "CSV"), the way the service would
        /// for a user. Needs a workspace on Premium or Fabric capacity.
        /// </summary>
        public async Task<byte[]> ExportReportAsync(
            string workspaceId,
            string reportId,
            string format,
            IReadOnlyDictionary<string, string?>? parameterValues = null,
            CancellationToken cancellationToken = default)
        {
            object body = parameterValues == null || parameterValues.Count == 0
                ? (object)new { format }
                : new
                {
                    format,
                    paginatedReportConfiguration = new
                    {
                        parameterValues = parameterValues.Select(p => new { name = p.Key, value = p.Value }).ToArray()
                    }
                };

            var export = await SendPowerBiAsync(HttpMethod.Post, $"/groups/{workspaceId}/reports/{reportId}/ExportTo", Json(body), cancellationToken).ConfigureAwait(false);
            string exportId = (string)export["id"]!;

            for (int attempt = 0; attempt < 100; attempt++)
            {
                await Task.Delay(TimeSpan.FromSeconds(3), cancellationToken).ConfigureAwait(false);

                var status = await SendPowerBiAsync(HttpMethod.Get, $"/groups/{workspaceId}/reports/{reportId}/exports/{exportId}", null, cancellationToken).ConfigureAwait(false);
                string state = (string)status["status"]!;

                if (state == "Succeeded")
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, $"{PowerBiApi}/groups/{workspaceId}/reports/{reportId}/exports/{exportId}/file"))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await getPowerBiToken(cancellationToken).ConfigureAwait(false));
                        using (var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                        {
                            await EnsureSuccessAsync(response).ConfigureAwait(false);
                            return await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                        }
                    }
                }

                if (state == "Failed")
                    throw new PowerBiException("Power BI couldn't export the report: " + DescribeError(status["error"]));
            }

            throw new PowerBiException("Power BI is still exporting the report after five minutes.");
        }

        public Task DeleteReportAsync(string workspaceId, string reportId, CancellationToken cancellationToken = default) =>
            SendPowerBiAsync(HttpMethod.Delete, $"/groups/{workspaceId}/reports/{reportId}", null, cancellationToken);

        public async Task DeleteConnectionAsync(string connectionId, CancellationToken cancellationToken = default) =>
            await SendFabricAsync(HttpMethod.Delete, $"/connections/{connectionId}", null, cancellationToken).ConfigureAwait(false);

        private async Task<JObject> SendPowerBiAsync(HttpMethod method, string path, HttpContent? content, CancellationToken cancellationToken)
        {
            string token = await getPowerBiToken(cancellationToken).ConfigureAwait(false);
            return await SendAsync(method, PowerBiApi + path, token, content, cancellationToken).ConfigureAwait(false);
        }

        private async Task<JObject> SendFabricAsync(HttpMethod method, string path, HttpContent? content, CancellationToken cancellationToken)
        {
            string? token = getConnectionsToken == null
                ? null
                : await getConnectionsToken(cancellationToken).ConfigureAwait(false);

            if (token == null)
                throw new PowerBiException("Flank hasn't been given access to Power BI connections.");

            return await SendAsync(method, FabricApi + path, token, content, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<JObject> SendAsync(HttpMethod method, string url, string token, HttpContent? content, CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(method, url) { Content = content })
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using (var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    await EnsureSuccessAsync(response).ConfigureAwait(false);

                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return string.IsNullOrWhiteSpace(body) ? new JObject() : JObject.Parse(body);
                }
            }
        }

        private static async Task EnsureSuccessAsync(HttpResponseMessage response)
        {
            if (response.IsSuccessStatusCode)
                return;

            string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            string detail = body;

            try
            {
                var json = JObject.Parse(body);
                detail = DescribeError(json["error"] ?? json);
            }
            catch (JsonException)
            {
            }

            throw new PowerBiException(
                $"Power BI returned {(int)response.StatusCode} {response.ReasonPhrase}" +
                (string.IsNullOrWhiteSpace(detail) ? "." : ": " + detail));
        }

        private static string DescribeError(JToken? error)
        {
            if (error == null)
                return "(no details)";

            string? message = (string?)error["message"];
            string? code = (string?)error["code"] ?? (string?)error["errorCode"];
            string? pbiCode = (string?)error["pbi.error"]?["code"];

            return string.Join(" ", new[] { pbiCode ?? code, message }.Where(s => !string.IsNullOrEmpty(s)))
                is var text && text.Length > 0 ? text : error.ToString(Formatting.None);
        }

        private static StringContent Json(object body) =>
            new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");
    }
}
