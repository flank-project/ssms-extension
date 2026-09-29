using System;
using System.Net;
using System.ServiceModel;
using ServiceReference;

namespace Flank.Ssrs
{
    public class SsrsClient
    {
        public async System.Threading.Tasks.Task TestConnectionAsync()
        {
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);

            binding.Security.Transport.ClientCredentialType =
                HttpClientCredentialType.Ntlm;

            var endpoint = new EndpointAddress(
                "http://localhost/ReportServer/ReportService2010.asmx");

            var client = new ReportingService2010SoapClient(binding, endpoint);

            new NetworkCredential(
                "user",
                "password",
                Environment.MachineName);

            var response = await client.ListChildrenAsync(
                null,
                "/",
                false);

            foreach (var item in response.CatalogItems)
            {
                Console.WriteLine($"{item.TypeName}: {item.Name}");
            }
        }
    }
}