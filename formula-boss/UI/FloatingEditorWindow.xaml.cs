using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Xml;

using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using ICSharpCode.AvalonEdit.Indentation.CSharp;

namespace FormulaBoss.UI;

public partial class FloatingEditorWindow
{
    private readonly DispatcherTimer _saveTimer;
    private readonly EditorSettings _settings;
    private CompletionWindow? _completionWindow;
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
        if (stream == null)
        {
            return;
        }

        using var reader = new XmlTextReader(stream);
        FormulaEditor.SyntaxHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);
    }

    private void OnTextEntering(object sender, TextCompositionEventArgs e)
    {
        if (e.Text.Length != 1)
        {
            return;
        }

        // Let completion window decide whether to accept/dismiss on this keystroke
        if (_completionWindow != null)
        {
            if (!char.IsLetterOrDigit(e.Text[0]) && e.Text[0] != '.')
            {
                _completionWindow.CompletionList.RequestInsertion(e);
            }
        }

        if (!e.Handled && EditorBehaviors.TrySkipClosingChar(FormulaEditor, e.Text[0]))
        {
            e.Handled = true;
        }
    }

    private void OnTextEntered(object sender, TextCompositionEventArgs e)
    {
        if (e.Text.Length != 1)
        {
            return;
        }

        var ch = e.Text[0];
        if (ch == '{')
        {
            EditorBehaviors.DeIndentOpenBrace(FormulaEditor);
        }

        EditorBehaviors.AutoInsertClosingChar(FormulaEditor, ch);

        // Trigger completion on '.' or any letter
        if (ch == '.' || char.IsLetter(ch))
        {
            ShowCompletion();
        }
    }

    private void ShowCompletion()
    {
        // Don't open a second window
        if (_completionWindow != null)
        {
            return;
        }

        var textUpToCaret = FormulaEditor.Document.GetText(0, FormulaEditor.CaretOffset);
        var items = CompletionProvider.GetCompletions(textUpToCaret);
        if (items.Count == 0)
        {
            return;
        }

        var wordLength = CompletionProvider.GetWordLength(textUpToCaret);

        // Pre-filter: don't show if the typed prefix doesn't match any item
        if (wordLength > 0)
        {
            var prefix = textUpToCaret[^wordLength..];
            if (!items.Any(i => i.Text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }
        }

        _completionWindow = new CompletionWindow(FormulaEditor.TextArea)
        {
            MinWidth = 250,
            CloseWhenCaretAtBeginning = true
        };

        // Enable filtering mode (default only highlights best match, doesn't hide non-matches)
        _completionWindow.CompletionList.IsFiltering = true;

        if (wordLength > 0)
        {
            _completionWindow.StartOffset = FormulaEditor.CaretOffset - wordLength;
        }

        foreach (var item in items)
        {
            _completionWindow.CompletionList.CompletionData.Add(item);
        }

        _completionWindow.Show();
        _completionWindow.Closed += (_, _) => _completionWindow = null;

        // Force initial filter - the document change that triggered us already
        // happened before the window was hooked into document events
        if (wordLength > 0)
        {
            _completionWindow.CompletionList.SelectItem(textUpToCaret[^wordLength..]);
        }
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
