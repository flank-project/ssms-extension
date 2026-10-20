using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Flank.PowerBi
{
    /// <summary>
    /// Signs the user in to Power BI as themselves (delegated), through the
    /// Flank app registration. Tokens are cached per Windows user, encrypted
    /// with DPAPI, in %LOCALAPPDATA%\Flank, so after the first sign-in the
    /// next ones are silent: in SSMS and in Flank.TestPbiApp alike.
    /// </summary>
    public sealed class PowerBiSignIn
    {
        // "Flank (dev)": a public client, single tenant (the test tenant).
        // TODO before release: a multi-tenant registration in Flank's own tenant.
        public const string ClientId = "80624e1f-0487-4fd3-9ea6-828e89464990";
        public const string TenantId = "da8d720b-2bc9-468d-b78d-08d46f3d3628";

        private const string CacheFileName = "powerbi.msalcache";

        /// <summary>The Power BI REST API: workspaces, reports, imports, data sources.</summary>
        public static readonly string[] PowerBiScopes =
        {
            "https://analysis.windows.net/powerbi/api/Report.ReadWrite.All",
            "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
            "https://analysis.windows.net/powerbi/api/Workspace.Read.All"
        };

        /// <summary>The Fabric REST API's shareable cloud connections.</summary>
        public static readonly string[] ConnectionScopes =
        {
            "https://api.fabric.microsoft.com/Connection.Read.All"
        };

        private readonly IPublicClientApplication app;

        private PowerBiSignIn(IPublicClientApplication app)
        {
            this.app = app;
        }

        public static string CacheDirectory =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Flank");

        public static async Task<PowerBiSignIn> CreateAsync()
        {
            var app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                .WithRedirectUri("http://localhost")
                .Build();

            // On Windows the cache file is encrypted with DPAPI for the current user.
            var storage = new StorageCreationPropertiesBuilder(CacheFileName, CacheDirectory).Build();
            var cache = await MsalCacheHelper.CreateAsync(storage).ConfigureAwait(false);
            cache.RegisterCache(app.UserTokenCache);

            return new PowerBiSignIn(app);
        }

        /// <summary>A token from the cache, or null if the user has to sign in.</summary>
        public async Task<PowerBiToken?> GetTokenSilentAsync(
            string[]? scopes = null,
            CancellationToken cancellationToken = default)
        {
            var account = (await app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            if (account == null)
                return null;

            try
            {
                var result = await app.AcquireTokenSilent(scopes ?? PowerBiScopes, account)
                    .ExecuteAsync(cancellationToken)
                    .ConfigureAwait(false);

                return new PowerBiToken(result.AccessToken, result.Account.Username);
            }
            catch (MsalUiRequiredException)
            {
                return null;
            }
        }

        /// <summary>
        /// Opens the system browser for the user to sign in. (On .NET Framework
        /// MSAL would otherwise use the old embedded Internet Explorer control.)
        /// </summary>
        /// <param name="parentWindow">The window to centre the sign-in on, or IntPtr.Zero.</param>
        /// <param name="scopes">What to ask for; Power BI by default.</param>
        public async Task<PowerBiToken> SignInInteractiveAsync(
            IntPtr parentWindow,
            CancellationToken cancellationToken = default,
            string[]? scopes = null)
        {
            var request = app.AcquireTokenInteractive(scopes ?? PowerBiScopes)
                .WithUseEmbeddedWebView(false)
                .WithPrompt(Prompt.SelectAccount);

            if (parentWindow != IntPtr.Zero)
                request = request.WithParentActivityOrWindow(parentWindow);

            var result = await request.ExecuteAsync(cancellationToken).ConfigureAwait(false);
            return new PowerBiToken(result.AccessToken, result.Account.Username);
        }

        /// <summary>
        /// Signs in with a code the user enters at microsoft.com/devicelogin,
        /// in any browser on any device. For machines whose own browser can't
        /// run the sign-in page (e.g. Windows Server's IE enhanced security).
        /// </summary>
        /// <param name="showCode">Receives the message to show the user (URL and code).</param>
        public async Task<PowerBiToken> SignInWithDeviceCodeAsync(
            Action<string> showCode,
            CancellationToken cancellationToken = default,
            string[]? scopes = null)
        {
            var result = await app.AcquireTokenWithDeviceCode(scopes ?? PowerBiScopes, code =>
                {
                    showCode(code.Message);
                    return Task.CompletedTask;
                })
                .ExecuteAsync(cancellationToken)
                .ConfigureAwait(false);

            return new PowerBiToken(result.AccessToken, result.Account.Username);
        }

        /// <summary>Forgets the cached sign-in (e.g. to switch account).</summary>
        public async Task SignOutAsync()
        {
            foreach (var account in await app.GetAccountsAsync().ConfigureAwait(false))
                await app.RemoveAsync(account).ConfigureAwait(false);
        }

        /// <summary>Silent if possible, otherwise asks the user to sign in.</summary>
        public async Task<PowerBiToken> GetTokenAsync(
            IntPtr parentWindow,
            CancellationToken cancellationToken = default)
        {
            return await GetTokenSilentAsync(null, cancellationToken).ConfigureAwait(false)
                ?? await SignInInteractiveAsync(parentWindow, cancellationToken).ConfigureAwait(false);
        }
    }

    public sealed class PowerBiToken
    {
        public PowerBiToken(string accessToken, string username)
        {
            AccessToken = accessToken;
            Username = username;
        }

        public string AccessToken { get; }

        public string Username { get; }
    }
}
