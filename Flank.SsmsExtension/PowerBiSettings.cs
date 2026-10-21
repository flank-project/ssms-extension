using Microsoft.Win32;
using System;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// The Power BI choices from the last report published, remembered per
    /// Windows user (HKCU\Software\Flank\PowerBi) to pre-fill the dialog next time.
    /// </summary>
    internal sealed class PowerBiSettings
    {
        private const string KeyPath = @"Software\Flank\PowerBi";

        public string? WorkspaceId { get; set; }

        public string? ConnectionId { get; set; }

        public bool UseSharedConnection { get; set; }

        public static PowerBiSettings Load()
        {
            var settings = new PowerBiSettings();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (key == null)
                        return settings;

                    settings.WorkspaceId = key.GetValue(nameof(WorkspaceId)) as string;
                    settings.ConnectionId = key.GetValue(nameof(ConnectionId)) as string;
                    settings.UseSharedConnection = key.GetValue(nameof(UseSharedConnection)) as int? == 1;
                }
            }
            catch (Exception ex)
            {
                // Remembering is a convenience; never block a report on it.
                System.Diagnostics.Debug.WriteLine($"Flank: could not read Power BI settings: {ex.Message}");
            }

            return settings;
        }

        public void Save()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    if (WorkspaceId != null) key.SetValue(nameof(WorkspaceId), WorkspaceId);
                    if (ConnectionId != null) key.SetValue(nameof(ConnectionId), ConnectionId);
                    key.SetValue(nameof(UseSharedConnection), UseSharedConnection ? 1 : 0, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Flank: could not save Power BI settings: {ex.Message}");
            }
        }
    }
}
