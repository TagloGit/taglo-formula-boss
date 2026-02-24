using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;

using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;


namespace FormulaBoss.UI;

public partial class FloatingEditorWindow
{
    private readonly DispatcherTimer _saveTimer;
    private readonly EditorSettings _settings;
    private readonly EditorBehaviorHandler _behaviorHandler;
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
        FormulaEditor.Options.IndentationSize = _settings.IndentSize;
        FormulaEditor.Options.ConvertTabsToSpaces = true;
        // Bracket matching highlight
        _ = new BracketHighlighter(FormulaEditor);

        // Real-time parse error squiggles
        _ = new ErrorHighlighter(FormulaEditor);

        // Editor behaviors (subscribes to TextEntering/TextEntered/PreviewKeyDown internally)
        _behaviorHandler = new EditorBehaviorHandler(FormulaEditor)
        {
            CompletionRequested = _ => ShowCompletion(),
            CompletionCloseRequested = _ => _completionWindow?.Close(),
            ForceCompletionRequested = () =>
            {
                _completionWindow?.Close();
                _completionWindow = null;
                ShowCompletion();
            },
            FormulaApplyRequested = text =>
            {
                FormulaApplied?.Invoke(this, text);
                Hide();
            },
            IsCompletionListEmpty = () =>
                _completionWindow != null && _completionWindow.CompletionList.ListBox.Items.Count == 0,
            IsCompletionWindowOpen = () => _completionWindow != null
        };

        // Track size changes with debounced save
        _saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _saveTimer.Tick += SaveTimer_Tick;
        SizeChanged += OnSizeChanged;

        // Focus the editor when window opens
        Activated += (_, _) =>
        {
            FormulaEditor.Focus();
            var text = FormulaEditor.Text.TrimEnd();
            if (text is "" or "=")
            {
                FormulaEditor.CaretOffset = FormulaEditor.Text.Length;
            }
            else
            {
                FormulaEditor.SelectAll();
            }
        };
    }

    /// <summary>
    ///     Workbook metadata (table names, named ranges, column headers) captured on
    ///     the Excel thread when the editor opens. Used for context-aware completions.
    /// </summary>
    public WorkbookMetadata? Metadata { get; set; }

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

    private void ShowCompletion()
    {
        // Don't open a second window
        if (_completionWindow != null)
        {
            return;
        }

        var textUpToCaret = FormulaEditor.Document.GetText(0, FormulaEditor.CaretOffset);
        var fullText = FormulaEditor.Text;
        var items = CompletionProvider.GetCompletions(textUpToCaret, fullText, Metadata, out var isBracketContext);
        if (items.Count == 0)
        {
            return;
        }

        _behaviorHandler.IsBracketContext = isBracketContext;
        var wordLength = CompletionProvider.GetWordLength(textUpToCaret, isBracketContext);

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
            MaxWidth = 600,
            CloseWhenCaretAtBeginning = true,
            SizeToContent = SizeToContent.WidthAndHeight
        };

        // Enable filtering mode (default only highlights best match, doesn't hide non-matches)
        _completionWindow.CompletionList.IsFiltering = true;

        // Style the completion list
        var listBox = _completionWindow.CompletionList.ListBox;
        ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
        listBox.Resources[SystemColors.HighlightBrushKey] = new SolidColorBrush(Color.FromRgb(180, 205, 235));
        listBox.Resources[SystemColors.HighlightTextBrushKey] = Brushes.Black;

        if (wordLength > 0)
        {
            _completionWindow.StartOffset = FormulaEditor.CaretOffset - wordLength;
        }

        foreach (var item in items)
        {
            _completionWindow.CompletionList.CompletionData.Add(item);
        }

        // Force initial filter - the document change that triggered us already
        // happened before the window was hooked into document events
        if (wordLength > 0)
        {
            _completionWindow.CompletionList.SelectItem(textUpToCaret[^wordLength..]);
        }

        _completionWindow.Show();
        _completionWindow.Closed += (_, _) =>
        {
            _completionWindow = null;
            _behaviorHandler.IsBracketContext = false;
        };
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
