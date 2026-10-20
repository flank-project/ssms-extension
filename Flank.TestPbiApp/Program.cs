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
    ///   Flank.TestPbiApp --connections  sign in with a code for listing connections, then list them
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

                // Grants the connections permission (listing Gateway / Cloud
                // Connections) with a code, then lists them.
                if (args.Length > 0 && args[0] == "--connections")
                {
                    var granted = await signIn.SignInWithDeviceCodeAsync(Console.WriteLine, scopes: PowerBiSignIn.ConnectionScopes);
                    Console.WriteLine($"Signed in for connections: {granted.Username} (object id {ObjectId(granted.AccessToken)})");

                    var client = new PowerBiClient(
                        async c => (await signIn.GetTokenSilentAsync(null, c))!.AccessToken,
                        async c => (await signIn.GetTokenSilentAsync(PowerBiSignIn.ConnectionScopes, c))?.AccessToken);

                    foreach (var connection in await client.GetConnectionsAsync())
                        Console.WriteLine($"  {connection.DisplayName}  [{connection.Kind}] {connection.Server};{connection.Database}");

                    Console.WriteLine("Connections API (silent): " +
                        (await signIn.GetTokenSilentAsync(PowerBiSignIn.ConnectionScopes) != null ? "OK" : "FAILED"));
                    return 0;
                }
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

                string connections;
                try
                {
                    connections = await signIn.GetTokenSilentAsync(PowerBiSignIn.ConnectionScopes) != null
                        ? "OK" : "needs sign-in";
                }
                catch (Exception ex)
                {
                    connections = ex.Message;
                }
                Console.WriteLine($"Connections API (silent): {connections}");
                Console.WriteLine($"Token cache: {PowerBiSignIn.CacheDirectory}");
                return again != null ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed: {ex.Message}");
                return 1;
            }
        }

        /// <summary>The signed-in user's Entra object id (the token's oid claim), e.g. to share a connection with them.</summary>
        private static string ObjectId(string accessToken)
        {
            string payload = accessToken.Split('.')[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '=');
            string json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(payload));
            var match = System.Text.RegularExpressions.Regex.Match(json, "\"oid\"\\s*:\\s*\"([^\"]+)\"");
            return match.Success ? match.Groups[1].Value : "(unknown)";
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
