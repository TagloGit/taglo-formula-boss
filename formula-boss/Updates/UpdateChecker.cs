using System.Net.Http;
using System.Reflection;
using System.Text.Json;

namespace FormulaBoss.Updates;

/// <summary>
///     Checks for newer releases on GitHub and exposes result for ribbon notification.
///     All failures are silent — network errors are logged but never shown to the user.
/// </summary>
internal static class UpdateChecker
{
    private static readonly Uri LatestReleaseUri =
        new("https://api.github.com/repos/TagloGit/taglo-formula-boss/releases/latest");

    /// <summary>
    ///     The newer version string (e.g. "0.2.0") if an update is available, otherwise null.
    /// </summary>
    public static string? NewVersionAvailable { get; private set; }

    /// <summary>
    ///     The URL of the GitHub release page for the newer version, otherwise null.
    /// </summary>
    public static string? ReleaseUrl { get; private set; }

    /// <summary>
    ///     Raised when a newer version is detected. Subscribers should call
    ///     <see cref="IRibbonUI.InvalidateControl" /> to refresh the ribbon.
    /// </summary>
    public static event Action? UpdateAvailable;

    /// <summary>
    ///     Fire-and-forget async check. Call from <see cref="AddIn.AutoOpen" /> without awaiting.
    /// </summary>
    public static async void CheckForUpdateAsync()
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            client.DefaultRequestHeaders.UserAgent.ParseAdd($"FormulaBoss/{version}");

            var json = await client.GetStringAsync(LatestReleaseUri);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var tagName = root.GetProperty("tag_name").GetString();
            var htmlUrl = root.GetProperty("html_url").GetString();

            if (tagName == null || htmlUrl == null)
            {
                return;
            }

            var remoteVersion = ParseVersion(tagName);
            if (remoteVersion == null)
            {
                return;
            }

            var currentVersion = version ?? new Version(0, 0, 0);
            if (remoteVersion > currentVersion)
            {
                NewVersionAvailable = remoteVersion.ToString();
                ReleaseUrl = htmlUrl;
                Logger.Info($"Update available: v{NewVersionAvailable} (current: v{currentVersion})");
                UpdateAvailable?.Invoke();
            }
            else
            {
                Logger.Info($"No update available (current: v{currentVersion}, latest: v{remoteVersion})");
            }
        }
        catch (Exception ex)
        {
            Logger.Info($"Update check failed (silent): {ex.Message}");
        }
    }

    /// <summary>
    ///     Parses a version string like "v0.2.0" or "0.2.0" into a <see cref="Version" />.
    ///     Returns null if parsing fails.
    /// </summary>
    internal static Version? ParseVersion(string tag)
    {
        var cleaned = tag.TrimStart('v', 'V');
        return Version.TryParse(cleaned, out var version) ? version : null;
    }
}
