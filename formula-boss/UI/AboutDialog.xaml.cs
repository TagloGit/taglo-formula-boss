using System.Diagnostics;
using System.Reflection;
using System.Windows;

namespace FormulaBoss.UI;

public partial class AboutDialog
{
    private const string GitHubUrl = "https://github.com/TagloGit/taglo-formula-boss";
    private const string ReleasesUrl = "https://github.com/TagloGit/taglo-formula-boss/releases/latest";

    public AboutDialog()
    {
        InitializeComponent();
        LoadLogo();

        var version = Assembly.GetExecutingAssembly().GetName().Version;
        VersionText.Text = version != null ? $"v{version.Major}.{version.Minor}.{version.Build}" : "";
    }

    private void LoadLogo()
    {
        var assembly = typeof(AboutDialog).Assembly;
        using var stream = assembly.GetManifestResourceStream("FormulaBoss.Resources.logo32.png");
        if (stream == null)
        {
            return;
        }

        var bitmap = new System.Windows.Media.Imaging.BitmapImage();
        bitmap.BeginInit();
        bitmap.StreamSource = stream;
        bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
        bitmap.EndInit();
        bitmap.Freeze();
        LogoImage.Source = bitmap;
    }

    private void OnGitHub(object sender, RoutedEventArgs e) =>
        Process.Start(new ProcessStartInfo(GitHubUrl) { UseShellExecute = true });

    private void OnReleaseNotes(object sender, RoutedEventArgs e) =>
        Process.Start(new ProcessStartInfo(ReleasesUrl) { UseShellExecute = true });

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
