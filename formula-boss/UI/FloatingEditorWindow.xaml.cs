using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;

namespace FormulaBoss.UI;

public partial class FloatingEditorWindow
{
    private readonly DispatcherTimer _saveTimer;
    private readonly EditorSettings _settings;
    private bool _sizeChanged;

    public FloatingEditorWindow()
    {
        _settings = EditorSettings.Load();

        InitializeComponent();

        // Apply saved size
        Width = _settings.Width;
        Height = _settings.Height;

        // Track size changes with debounced save
        _saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _saveTimer.Tick += SaveTimer_Tick;
        SizeChanged += OnSizeChanged;

        // Focus the text box when window opens
        Activated += (_, _) =>
        {
            FormulaTextBox.Focus();
            FormulaTextBox.SelectAll();
        };
    }

    public string FormulaText
    {
        get => FormulaTextBox.Text;
        set => FormulaTextBox.Text = value;
    }

    public event EventHandler<string>? FormulaApplied;

    private void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        _sizeChanged = true;
        _saveTimer.Stop();
        _saveTimer.Start();
    }

    private void SaveTimer_Tick(object? sender, EventArgs e)
    {
        _saveTimer.Stop();
        if (_sizeChanged)
        {
            _settings.Width = Width;
            _settings.Height = Height;
            _settings.Save();
            _sizeChanged = false;
        }
    }

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        FormulaApplied?.Invoke(this, FormulaText);
        Hide();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => Hide();

    protected override void OnClosing(CancelEventArgs e)
    {
        // Don't actually close, just hide - we reuse the window
        e.Cancel = true;
        Hide();
    }
}
