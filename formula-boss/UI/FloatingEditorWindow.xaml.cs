using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Xml;

using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using ICSharpCode.AvalonEdit.Indentation.CSharp;

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

        // Load syntax highlighting
        LoadSyntaxHighlighting();

        // Enable auto-indentation
        FormulaEditor.Options.IndentationSize = 4;
        FormulaEditor.Options.ConvertTabsToSpaces = true;
        FormulaEditor.TextArea.IndentationStrategy =
            new CSharpIndentationStrategy(FormulaEditor.Options);

        // Editor behaviors
        FormulaEditor.TextArea.TextEntering += OnTextEntering;
        FormulaEditor.TextArea.TextEntered += OnTextEntered;
        FormulaEditor.PreviewKeyDown += OnPreviewKeyDown;

        // Track size changes with debounced save
        _saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _saveTimer.Tick += SaveTimer_Tick;
        SizeChanged += OnSizeChanged;

        // Focus the editor when window opens
        Activated += (_, _) =>
        {
            FormulaEditor.Focus();
            FormulaEditor.SelectAll();
        };
    }

    public string FormulaText
    {
        get => FormulaEditor.Text;
        set => FormulaEditor.Text = value;
    }

    public event EventHandler<string>? FormulaApplied;

    private void LoadSyntaxHighlighting()
    {
        var assembly = typeof(FloatingEditorWindow).Assembly;
        using var stream = assembly.GetManifestResourceStream("FormulaBoss.UI.FormulaBossSyntax.xshd");
        if (stream == null) return;

        using var reader = new XmlTextReader(stream);
        FormulaEditor.SyntaxHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
    }

    private void OnTextEntering(object sender, TextCompositionEventArgs e)
    {
        if (e.Text.Length == 1 && EditorBehaviors.TrySkipClosingChar(FormulaEditor, e.Text[0]))
            e.Handled = true;
    }

    private void OnTextEntered(object sender, TextCompositionEventArgs e)
    {
        if (e.Text.Length != 1) return;

        var ch = e.Text[0];
        if (ch == '{') EditorBehaviors.DeIndentOpenBrace(FormulaEditor);
        EditorBehaviors.AutoInsertClosingChar(FormulaEditor, ch);
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && e.KeyboardDevice.Modifiers == ModifierKeys.None
            && EditorBehaviors.TryExpandBraceBlock(FormulaEditor))
        {
            e.Handled = true;
            return;
        }

        if (e is { Key: Key.Enter, KeyboardDevice.Modifiers: ModifierKeys.Control })
        {
            e.Handled = true;
            FormulaApplied?.Invoke(this, FormulaText);
            Hide();
        }
    }

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
