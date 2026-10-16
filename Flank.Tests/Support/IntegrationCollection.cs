using Flank.Ssrs.ReportService2010;
using System.Net;

namespace Flank.Tests.Support;

/// <summary>
/// Integration tests run one at a time (they drive Excel and SSRS) and share
/// a clean SSRS test folder.
/// </summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public class IntegrationCollection : ICollectionFixture<SsrsTestFolder>, ICollectionFixture<PowerBiTestWorkspace>
{
    public const string Name = "Integration";
}

/// <summary>Empties the SSRS test folder once per run.</summary>
public sealed class SsrsTestFolder
{
    private bool prepared;

    public string Path => TestSettings.ReportFolder;

    /// <summary>Called lazily so runs that skip SSRS tests don't touch the server.</summary>
    public void EnsureClean()
    {
        if (prepared)
            return;

        var rs = new ReportingService2010
        {
            Url = TestSettings.ReportServer.TrimEnd('/') + "/ReportService2010.asmx",
            Credentials = CredentialCache.DefaultCredentials,
        };

        string name = Path.TrimStart('/');

        if (rs.ListChildren("/", false).Any(i => i.Path == Path))
            rs.DeleteItem(Path);

        rs.CreateFolder(name, "/", null);
        prepared = true;
    }
}
