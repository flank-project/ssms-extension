using ServiceReference;
using System.ServiceModel;
using System;

namespace Flank.Ssrs;

public class SsrsClient
{
    public async System.Threading.Tasks.Task TestConnectionAsync()
    {
        var client = new ReportingService2010SoapClient(
            ReportingService2010SoapClient.EndpointConfiguration.ReportingService2010Soap);

        client.Endpoint.Address = new EndpointAddress(
            "http://localhost/ReportServer/ReportService2010.asmx");

        client.ClientCredentials.Windows.ClientCredential =
            System.Net.CredentialCache.DefaultNetworkCredentials;

        var response = await client.ListChildrenAsync(
            null,   // TrustedUserHeader
            "/",    // ItemPath
            false   // Recursive
        );

        foreach (var item in response.CatalogItems)
        {
            Console.WriteLine($"{item.TypeName}: {item.Name}");
        }
    }
}
