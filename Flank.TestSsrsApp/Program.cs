using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flank.TestSsrsApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connection = new SqlConnection(
                "Server=;" +
                "Database=;" +
                "User ID=;" +
                "Password=;"
            );

            connection.Open();

            var client = new Flank.Ssrs.SsrsClient(
                "http://localhost/ReportServer",
                "http://localhost/Reports");

            Console.WriteLine("FOLDERS:");

            foreach (var folder in client.GetFolders())
            {
                Console.WriteLine(folder);
            }

            Console.WriteLine();

            Console.WriteLine("DATA SOURCES:");

            foreach (var dataSource in client.GetSharedDataSources())
            {
                Console.WriteLine(dataSource);
            }

            string reportUrl = client.CreateReportFromStoredProcedure(
                "dbo.becky_driver_example3",
                connection,
                "/HardcodedTest",
                "/TestFolder",
                "becky_driver_example3 2");

            Console.WriteLine(reportUrl);

            Console.ReadLine();
        }
    }
}
