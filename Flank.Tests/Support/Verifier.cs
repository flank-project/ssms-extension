using System.Diagnostics;

namespace Flank.Tests.Support;

/// <summary>
/// Runs the checker scripts in tests\ (also used by the SSMS smoke tests), so
/// there's one definition of "the workbook refreshes" and "the report renders".
/// </summary>
internal static class Verifier
{
    public static string VerifyExcel(string workbookPath, int minRows = 1) =>
        Run("verify-excel.ps1", "-Path", workbookPath, "-MinRows", minRows.ToString());

    /// <param name="parameters">Report parameters as "name=value;name2=value2".</param>
    public static string VerifySsrs(string report, string parameters = "", int minRows = 1) =>
        Run("verify-ssrs.ps1",
            "-Report", report,
            "-Parameters", parameters,
            "-ReportServer", TestSettings.ReportServer,
            "-MinRows", minRows.ToString());

    /// <param name="report">Report name in the test workspace, or the link Flank shows.</param>
    /// <param name="parameters">Report parameters as "name=value;name2=value2".</param>
    public static string VerifyPowerBi(string report, string parameters = "", int minRows = 1) =>
        Run("verify-pbi.ps1",
            "-Report", report,
            "-Parameters", parameters,
            "-MinRows", minRows.ToString());

    private static string Run(string script, params string[] args)
    {
        string scriptPath = Path.Combine(RepoRoot.Path, "tests", script);
        string quoted = string.Join(" ", args.Select(a => a.StartsWith("-") ? a : $"\"{a}\""));

        var start = new ProcessStartInfo(
            "powershell.exe",
            $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\" {quoted}")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        using var process = Process.Start(start)!;
        var stderr = process.StandardError.ReadToEndAsync();
        string stdout = process.StandardOutput.ReadToEnd();
        process.WaitForExit();

        string output = (stdout + stderr.Result).Trim();
        Assert.True(process.ExitCode == 0, $"{script} failed (exit {process.ExitCode}):\n{output}");
        return output;
    }
}

internal static class RepoRoot
{
    public static string Path { get; } = Find();

    private static string Find()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
        {
            if (File.Exists(System.IO.Path.Combine(dir.FullName, "ssms-extension.slnx")))
                return dir.FullName;
        }

        throw new InvalidOperationException("Could not find the repo root (ssms-extension.slnx).");
    }
}
