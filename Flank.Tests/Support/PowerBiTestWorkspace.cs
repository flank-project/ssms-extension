using Flank.PowerBi;
using Newtonsoft.Json.Linq;
using System.Data.SqlClient;
using System.Net.Http;

namespace Flank.Tests.Support;

/// <summary>
/// The Power BI workspace the tests publish to, signed in as the test service
/// principal from FLANK_TEST_PBI ("TenantId=...;ClientId=...;ClientSecret=...;Workspace=...").
/// Test reports are named "Test - ..." and deleted at the start of each run;
/// shared connections the tests create are deleted by the tests themselves.
/// </summary>
public sealed class PowerBiTestWorkspace
{
    public const string ReportPrefix = "Test - ";

    private static readonly HttpClient Http = new();

    private readonly Dictionary<string, (string Token, DateTime Expires)> tokens = new();
    private Dictionary<string, string>? settings;
    private PowerBiClient? client;
    private string? workspaceId;

    public PowerBiClient Client => client ?? throw new InvalidOperationException("Call EnsureReadyAsync first.");

    public string WorkspaceId => workspaceId ?? throw new InvalidOperationException("Call EnsureReadyAsync first.");

    /// <summary>The SQL login the tests save on reports and connections (from FLANK_TEST_SQL).</summary>
    public static (string Username, string Password) SqlLogin
    {
        get
        {
            var builder = new SqlConnectionStringBuilder(TestSettings.SqlConnectionString);
            return (builder.UserID, builder.Password);
        }
    }

    /// <summary>Signs in, finds the workspace and removes earlier test reports. Skips the test without settings.</summary>
    public async Task EnsureReadyAsync()
    {
        settings ??= Parse(Environment.GetEnvironmentVariable("FLANK_TEST_PBI"));

        Assert.SkipWhen(
            settings == null,
            "Set FLANK_TEST_PBI (TenantId=...;ClientId=...;ClientSecret=...;Workspace=...) to run Power BI tests.");
        Assert.SkipWhen(
            TestSettings.SqlConnectionString == null,
            "Set FLANK_TEST_SQL to run Power BI tests.");

        if (client != null)
            return;

        var newClient = new PowerBiClient(
            _ => GetTokenAsync("https://analysis.windows.net/powerbi/api/.default"),
            async _ => await GetTokenAsync("https://api.fabric.microsoft.com/.default"));

        string name = settings!["workspace"];
        var workspace = (await newClient.GetWorkspacesAsync()).FirstOrDefault(w => w.Name == name)
            ?? throw new InvalidOperationException($"The test service principal can't see workspace '{name}'.");

        foreach (var report in await newClient.GetReportsAsync(workspace.Id))
        {
            if (report.Name.StartsWith(ReportPrefix, StringComparison.Ordinal))
                await newClient.DeleteReportAsync(workspace.Id, report.Id);
        }

        workspaceId = workspace.Id;
        client = newClient;
    }

    private async Task<string> GetTokenAsync(string scope)
    {
        if (tokens.TryGetValue(scope, out var cached) && cached.Expires > DateTime.UtcNow.AddMinutes(5))
            return cached.Token;

        var response = await Http.PostAsync(
            $"https://login.microsoftonline.com/{settings!["tenantid"]}/oauth2/v2.0/token",
            new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = settings["clientid"],
                ["client_secret"] = settings["clientsecret"],
                ["scope"] = scope,
            }));

        string body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"Signing in as the test service principal failed: {body}");

        var json = JObject.Parse(body);
        string token = (string)json["access_token"]!;
        tokens[scope] = (token, DateTime.UtcNow.AddSeconds((int)json["expires_in"]!));
        return token;
    }

    private static Dictionary<string, string>? Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (string pair in value!.Split(';'))
        {
            int equals = pair.IndexOf('=');
            if (equals > 0)
                result[pair.Substring(0, equals).Trim().ToLowerInvariant()] = pair.Substring(equals + 1).Trim();
        }

        return result.ContainsKey("tenantid") && result.ContainsKey("clientid") &&
               result.ContainsKey("clientsecret") && result.ContainsKey("workspace")
            ? result
            : null;
    }
}
