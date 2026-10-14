using System.Data.SqlClient;

namespace Flank.Tests.Support;

/// <summary>
/// Settings for integration tests, from environment variables with sandbox defaults.
/// </summary>
/// <remarks>
/// FLANK_TEST_SQL                 Connection string to the test database. Needed by tests that
///                                query metadata (SSRS reports); they're skipped without it.
/// FLANK_TEST_REPORT_SERVER       Default http://localhost/ReportServer
/// FLANK_TEST_REPORT_PORTAL       Default http://localhost/Reports
/// FLANK_TEST_REPORT_DATA_SOURCE  Shared data source pointing at the test database.
///                                Default /HardcodedTest
/// </remarks>
internal static class TestSettings
{
    private const string DefaultServer = "fomf-sandbox-ssdb.database.windows.net";
    private const string DefaultDatabase = "fomf-ss-sandbox-db";

    public static string? SqlConnectionString => Get("FLANK_TEST_SQL");

    /// <summary>Server for generated workbooks. Excel uses its own saved credentials.</summary>
    public static string Server =>
        SqlConnectionString is { } cs
            ? StripTcpPrefix(new SqlConnectionStringBuilder(cs).DataSource)
            : DefaultServer;

    public static string Database =>
        SqlConnectionString is { } cs
            ? new SqlConnectionStringBuilder(cs).InitialCatalog
            : DefaultDatabase;

    public static string ReportServer => Get("FLANK_TEST_REPORT_SERVER") ?? "http://localhost/ReportServer";

    public static string ReportPortal => Get("FLANK_TEST_REPORT_PORTAL") ?? "http://localhost/Reports";

    public static string ReportDataSource => Get("FLANK_TEST_REPORT_DATA_SOURCE") ?? "/HardcodedTest";

    /// <summary>SSRS folder the tests own; emptied at the start of each run.</summary>
    public const string ReportFolder = "/FlankTests";

    public static SqlConnection OpenSqlConnection()
    {
        Assert.SkipWhen(
            SqlConnectionString == null,
            "Set FLANK_TEST_SQL to a connection string for the test database to run this test.");

        var connection = new SqlConnection(SqlConnectionString);
        connection.Open();
        return connection;
    }

    private static string? Get(string name)
    {
        string? value = Environment.GetEnvironmentVariable(name);
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    // "tcp:host,1433" -> "host"
    private static string StripTcpPrefix(string dataSource)
    {
        if (dataSource.StartsWith("tcp:", StringComparison.OrdinalIgnoreCase))
            dataSource = dataSource.Substring(4);

        int comma = dataSource.IndexOf(',');
        return comma >= 0 ? dataSource.Substring(0, comma) : dataSource;
    }
}
