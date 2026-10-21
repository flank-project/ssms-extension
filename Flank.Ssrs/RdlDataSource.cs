using System;
using System.Security;

namespace Flank.Ssrs
{
    /// <summary>
    /// The data source a report definition (RDL) reads from: a shared data
    /// source on the SSRS server, or one embedded in the report (Power BI
    /// paginated reports, which have no shared data sources).
    /// </summary>
    public sealed class RdlDataSource
    {
        private readonly string? sharedPath;
        private readonly string? connectString;
        private readonly string? dataProvider;

        private RdlDataSource(
            string name,
            string? sharedPath,
            string? dataProvider,
            string? connectString)
        {
            Name = name;
            this.sharedPath = sharedPath;
            this.dataProvider = dataProvider;
            this.connectString = connectString;
        }

        /// <summary>The data source's name inside the report.</summary>
        public string Name { get; }

        /// <summary>A shared data source on the SSRS server, e.g. "/Data Sources/Main".</summary>
        public static RdlDataSource Shared(string path) =>
            new RdlDataSource("SharedDataSource", path, null, null);

        /// <summary>
        /// An Azure SQL database, embedded in the report. Credentials are set
        /// in the Power BI service, never in the report.
        /// </summary>
        public static RdlDataSource AzureSql(string server, string database) =>
            new RdlDataSource(
                "Database",
                null,
                "SQLAZURE",
                $"Data Source={server};Initial Catalog={database};Encrypt=True");

        /// <summary>A SQL Server database, embedded in the report (reached through an on-premises gateway).</summary>
        public static RdlDataSource SqlServer(string server, string database) =>
            new RdlDataSource(
                "Database",
                null,
                "SQL",
                $"Data Source={server};Initial Catalog={database}");

        /// <summary>
        /// The data source for a Power BI connection to this database: Azure SQL
        /// for an Azure host reached directly, SQL Server otherwise (gateways).
        /// </summary>
        public static RdlDataSource ForPowerBiConnection(string server, string database, bool throughGateway) =>
            !throughGateway && IsAzureSql(server)
                ? AzureSql(server, database)
                : SqlServer(server, database);

        /// <summary>An Azure SQL Database host (*.database.windows.net), with or without tcp: and a port.</summary>
        public static bool IsAzureSql(string server)
        {
            string host = server.Trim();

            if (host.StartsWith("tcp:", StringComparison.OrdinalIgnoreCase))
                host = host.Substring(4);

            int comma = host.IndexOf(',');
            if (comma >= 0)
                host = host.Substring(0, comma);

            return host.EndsWith(".database.windows.net", StringComparison.OrdinalIgnoreCase);
        }

        internal string ToXml()
        {
            if (sharedPath != null)
            {
                return $@"
    <DataSource Name=""{Escape(Name)}"">
      <DataSourceReference>{Escape(sharedPath)}</DataSourceReference>
    </DataSource>";
            }

            return $@"
    <DataSource Name=""{Escape(Name)}"">
      <ConnectionProperties>
        <DataProvider>{Escape(dataProvider!)}</DataProvider>
        <ConnectString>{Escape(connectString!)}</ConnectString>
      </ConnectionProperties>
      <rd:DataSourceID>{Guid.NewGuid()}</rd:DataSourceID>
    </DataSource>";
        }

        private static string Escape(string value) => SecurityElement.Escape(value) ?? "";
    }
}
