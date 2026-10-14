using Microsoft.Win32;
using System;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// The SSRS choices from the last report created, remembered per Windows
    /// user (HKCU\Software\Flank\Ssrs) to pre-fill the dialogs next time.
    /// </summary>
    internal sealed class SsrsSettings
    {
        private const string KeyPath = @"Software\Flank\Ssrs";

        public string ReportServerUrl { get; set; } = "http://localhost/ReportServer";
        public string ReportPortalUrl { get; set; } = "http://localhost/Reports";
        public string? Folder { get; set; }
        public string? DataSource { get; set; }

        public static SsrsSettings Load()
        {
            var settings = new SsrsSettings();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (key == null)
                        return settings;

                    settings.ReportServerUrl = key.GetValue(nameof(ReportServerUrl)) as string ?? settings.ReportServerUrl;
                    settings.ReportPortalUrl = key.GetValue(nameof(ReportPortalUrl)) as string ?? settings.ReportPortalUrl;
                    settings.Folder = key.GetValue(nameof(Folder)) as string;
                    settings.DataSource = key.GetValue(nameof(DataSource)) as string;
                }
            }
            catch (Exception ex)
            {
                // Remembering is a convenience; never block a report on it.
                System.Diagnostics.Debug.WriteLine($"Flank: could not read SSRS settings: {ex.Message}");
            }

            return settings;
        }

        public void Save()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    key.SetValue(nameof(ReportServerUrl), ReportServerUrl);
                    key.SetValue(nameof(ReportPortalUrl), ReportPortalUrl);
                    if (Folder != null) key.SetValue(nameof(Folder), Folder);
                    if (DataSource != null) key.SetValue(nameof(DataSource), DataSource);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Flank: could not save SSRS settings: {ex.Message}");
            }
        }
    }
}
