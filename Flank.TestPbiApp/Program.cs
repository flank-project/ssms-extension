using Flank.PowerBi;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Flank.TestPbiApp
{
    /// <summary>
    /// Signs in to Power BI the way Flank does, filling the same token cache,
    /// so Flank in SSMS can sign in silently afterwards.
    ///   Flank.TestPbiApp          sign in (browser if needed), then list workspaces
    ///   Flank.TestPbiApp --device sign in with a code entered in any browser
    ///   Flank.TestPbiApp --check  silent only; exit code 1 if a sign-in is needed
    /// </summary>
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            bool checkOnly = args.Length > 0 && args[0] == "--check";
            bool deviceCode = args.Length > 0 && args[0] == "--device";

            try
            {
                var signIn = await PowerBiSignIn.CreateAsync();
                var token = await signIn.GetTokenSilentAsync();

                if (token != null)
                {
                    Console.WriteLine($"Cached sign-in: {token.Username}");
                }
                else if (checkOnly)
                {
                    Console.WriteLine("No cached sign-in: run Flank.TestPbiApp without --check to sign in.");
                    return 1;
                }
                else if (deviceCode)
                {
                    token = await signIn.SignInWithDeviceCodeAsync(Console.WriteLine);
                    Console.WriteLine($"Signed in: {token.Username}");
                }
                else
                {
                    Console.WriteLine("Opening the browser to sign in to Power BI...");
                    token = await signIn.SignInInteractiveAsync(IntPtr.Zero);
                    Console.WriteLine($"Signed in: {token.Username}");
                }

                await ListWorkspacesAsync(token.AccessToken);

                // The next sign-in (here, or Flank in SSMS) must come from the cache.
                var again = await signIn.GetTokenSilentAsync();
                Console.WriteLine(again != null ? "Silent sign-in: OK" : "Silent sign-in: FAILED");
                Console.WriteLine($"Token cache: {PowerBiSignIn.CacheDirectory}");
                return again != null ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed: {ex.Message}");
                return 1;
            }
        }

        private static async Task ListWorkspacesAsync(string accessToken)
        {
            using (var http = new HttpClient())
            {
                http.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await http.GetAsync("https://api.powerbi.com/v1.0/myorg/groups");
                var body = await response.Content.ReadAsStringAsync();

                Console.WriteLine(response.IsSuccessStatusCode
                    ? $"Workspaces: {body}"
                    : $"Listing workspaces failed: {(int)response.StatusCode} {body}");
            }
        }
    }
}
