using Flank.Ssrs;
using System;
using System.Data.SqlClient;

namespace Flank.TestSsrsApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var options = ParseArgs(args);

                var client = new SsrsClient(
                    GetRequired(options, "report-server"),
                    GetRequired(options, "report-portal"));

                string action = GetRequired(options, "action");

                if (action == "folders")
                {
                    foreach (var folder in client.GetFolders())
                    {
                        Console.WriteLine(folder);
                    }

                    return;
                }

                if (action == "data-sources")
                {
                    foreach (var dataSource in client.GetSharedDataSources())
                    {
                        Console.WriteLine(dataSource);
                    }

                    return;
                }

                if (action == "create")
                {
                    CreateReport(client, options);
                    return;
                }

                throw new ArgumentException(
                    $"Unknown action '{action}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR:");
                Console.Error.WriteLine(ex.Message);
                Environment.ExitCode = 1;
            }
        }

        private static void CreateReport(
            SsrsClient client,
            System.Collections.Generic.Dictionary<string, string> options)
        {
            string connectionString =
                GetRequired(options, "connection-string");

            string type =
                GetRequired(options, "type");

            string command =
                GetRequired(options, "command");

            string dataSource =
                GetRequired(options, "data-source");

            string folder =
                GetRequired(options, "folder");

            string reportName =
                GetRequired(options, "report-name");

            using (var connection =
                new SqlConnection(connectionString))
            {
                connection.Open();

                string reportUrl;

                if (type == "sql")
                {
                    reportUrl = client.CreateReport(
                        command,
                        connection,
                        dataSource,
                        folder,
                        reportName);
                }
                else if (type == "sproc")
                {
                    reportUrl =
                        client.CreateReportFromStoredProcedure(
                            command,
                            connection,
                            dataSource,
                            folder,
                            reportName);
                }
                else
                {
                    throw new ArgumentException(
                        "--type must be 'sql' or 'sproc'.");
                }

                Console.WriteLine("Report created:");
                Console.WriteLine(reportUrl);
            }
        }

        private static System.Collections.Generic.Dictionary<string, string>
            ParseArgs(string[] args)
        {
            var result =
                new System.Collections.Generic.Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                if (!arg.StartsWith("--"))
                {
                    throw new ArgumentException(
                        $"Unexpected argument '{arg}'.");
                }

                string key = arg.Substring(2);

                if (i + 1 >= args.Length ||
                    args[i + 1].StartsWith("--"))
                {
                    throw new ArgumentException(
                        $"Missing value for '--{key}'.");
                }

                result[key] = args[++i];
            }

            return result;
        }

        private static string GetRequired(
            System.Collections.Generic.Dictionary<string, string> options,
            string name)
        {
            string value;

            if (!options.TryGetValue(name, out value) ||
                string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException(
                    $"Missing required argument '--{name}'.");
            }

            return value;
        }
    }
}