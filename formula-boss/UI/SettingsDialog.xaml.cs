using System.Windows;

namespace FormulaBoss.UI;

public partial class SettingsDialog : Window
{
    public SettingsDialog(EditorSettings settings)
    {
        InitializeComponent();

        AnimationCombo.ItemsSource = Enum.GetValues<AnimationStyle>();
        AnimationCombo.SelectedItem = settings.AnimationStyle;

        IndentSizeBox.Text = settings.IndentSize.ToString();
    }

    public AnimationStyle SelectedAnimation =>
        AnimationCombo.SelectedItem is AnimationStyle style ? style : AnimationStyle.Chomp;

    public int SelectedIndentSize =>
        int.TryParse(IndentSizeBox.Text, out var size) && size is >= 1 and <= 8 ? size : 2;

    private void OnOk(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
