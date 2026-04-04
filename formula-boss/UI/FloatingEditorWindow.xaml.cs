using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Xml;

using FormulaBoss.UI.Completion;

using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;


namespace FormulaBoss.UI;

public partial class FloatingEditorWindow
{
    private readonly DispatcherTimer _saveTimer;
    private readonly EditorSettings _settings;
    private readonly EditorBehaviorHandler _behaviorHandler;
    private CompletionPopup? _completionPopup;
    private bool _sizeChanged;
    private RoslynWorkspaceManager? _workspaceManager;
    private RoslynCompletionProvider? _completionProvider;
    private SignatureHelpProvider? _signatureHelpProvider;
    private SignatureHelpPopup? _signatureHelpPopup;
    private CancellationTokenSource? _completionCts;
    private CancellationTokenSource? _signatureHelpCts;

    public FloatingEditorWindow()
    {
        _settings = EditorSettings.Load();

        InitializeComponent();

        // Load logo into the bottom bar
        LoadLogo();

        // Set version label
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        if (version != null)
        {
            VersionLabel.Text = $"v{version.Major}.{version.Minor}.{version.Build}";
        }

        // Apply saved size and font size
        Width = _settings.Width;
        Height = _settings.Height;
        FormulaEditor.FontSize = _settings.FontSize;
        FormulaEditor.WordWrap = _settings.WordWrap;

        // Ctrl+MouseWheel zoom
        FormulaEditor.PreviewMouseWheel += OnEditorPreviewMouseWheel;

        // Load syntax highlighting
        LoadSyntaxHighlighting();

        // Enable auto-indentation
        FormulaEditor.Options.IndentationSize = _settings.IndentSize;
        FormulaEditor.Options.ConvertTabsToSpaces = true;
        // Bracket matching highlight
        _ = new BracketHighlighter(FormulaEditor);

        // Real-time Roslyn error squiggles
        _ = new ErrorHighlighter(
            FormulaEditor,
            () => { EnsureWorkspace(); return _workspaceManager; },
            () => Metadata);

        // Editor behaviors (subscribes to TextEntering/TextEntered/PreviewKeyDown internally)
        _behaviorHandler = new EditorBehaviorHandler(FormulaEditor)
        {
            CompletionRequested = _ => ShowCompletion(),
            CompletionCloseRequested = _ => _completionPopup?.Close(),
            ForceCompletionRequested = () =>
            {
                _completionPopup?.Close();
                _completionPopup = null;
                ShowCompletion();
            },
            FormulaApplyRequested = text =>
            {
                DismissSignatureHelp();
                FormulaApplied?.Invoke(this, text);
                Hide();
            },
            IsCompletionListEmpty = () =>
                _completionPopup != null && _completionPopup.CompletionList.ListBox?.Items.Count == 0,
            IsCompletionWindowOpen = () => _completionPopup != null,
            SignatureHelpRequested = () => ShowSignatureHelp(),
            SignatureHelpDismissRequested = DismissSignatureHelp,
            IsSignatureHelpVisible = () => _signatureHelpPopup?.IsVisible == true,
            NextOverload = () => _signatureHelpPopup?.NextOverload() == true,
            PreviousOverload = () => _signatureHelpPopup?.PreviousOverload() == true
        };

        // Track size changes with debounced save
        _saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _saveTimer.Tick += SaveTimer_Tick;
        SizeChanged += OnSizeChanged;

        // Alt+W toggles word wrap
        var toggleWordWrapCommand = new RoutedCommand();
        CommandBindings.Add(new CommandBinding(toggleWordWrapCommand, (_, _) => ToggleWordWrap()));
        InputBindings.Add(new InputBinding(toggleWordWrapCommand, new KeyGesture(Key.W, ModifierKeys.Alt)));

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

        // Dismiss signature help when editor loses focus
        Deactivated += (_, _) => DismissSignatureHelp();
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

    /// <summary>
    ///     Updates the title bar to show the cell being edited.
    /// </summary>
    public void UpdateTitle(string? sheetName, string? address)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(address))
        {
            Title = "Formula Boss";
            return;
        }

        var cleanAddress = address.Replace("$", "");
        Title = $"Formula Boss \u2014 {sheetName}!{cleanAddress}";
    }

    public event EventHandler<string>? FormulaApplied;

    private void LoadLogo()
    {
        var assembly = typeof(FloatingEditorWindow).Assembly;
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

    private void EnsureWorkspace()
    {
        if (_workspaceManager != null)
        {
            return;
        }

        _workspaceManager = new RoslynWorkspaceManager();
        _completionProvider = new RoslynCompletionProvider(_workspaceManager);
        _signatureHelpProvider = new SignatureHelpProvider(_workspaceManager);

        // Fire-and-forget warm-up
        _ = _workspaceManager.WarmUpAsync();
    }

    private async void ShowCompletion()
    {
        // Don't open a second popup
        if (_completionPopup != null)
        {
            return;
        }

        EnsureWorkspace();

        // Cancel any in-flight completion request
        _completionCts?.Cancel();
        _completionCts = new CancellationTokenSource();
        var ct = _completionCts.Token;

        var textUpToCaret = FormulaEditor.Document.GetText(0, FormulaEditor.CaretOffset);
        var fullText = FormulaEditor.Text;

        IReadOnlyList<CompletionData> items;
        bool isBracketContext;

        try
        {
            (items, isBracketContext) = await _completionProvider!.GetCompletionsAsync(
                textUpToCaret, fullText, Metadata, ct);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (ct.IsCancellationRequested || items.Count == 0)
        {
            return;
        }

        // Don't open if another popup appeared while we were awaiting
        if (_completionPopup != null)
        {
            return;
        }

        _behaviorHandler.IsBracketContext = isBracketContext;
        var wordLength = CompletionHelpers.GetWordLength(textUpToCaret, isBracketContext);

        // Pre-filter: don't show if the typed prefix doesn't match any item
        if (wordLength > 0)
        {
            var prefix = textUpToCaret[^wordLength..];
            if (!items.Any(i => i.Text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }
        }

        _completionPopup = new CompletionPopup(FormulaEditor.TextArea)
        {
            CloseWhenCaretAtBeginning = true
        };

        if (wordLength > 0)
        {
            _completionPopup.StartOffset = FormulaEditor.CaretOffset - wordLength;
        }

        foreach (var item in items)
        {
            _completionPopup.CompletionList.CompletionData.Add(item);
        }

        // Force initial filter - the document change that triggered us already
        // happened before the popup was hooked into document events
        if (wordLength > 0)
        {
            _completionPopup.CompletionList.SelectItem(textUpToCaret[^wordLength..]);
        }

        _completionPopup.Closed += (_, _) =>
        {
            _completionPopup = null;
            _behaviorHandler.IsBracketContext = false;
        };
        _completionPopup.Show();
    }

    private async void ShowSignatureHelp()
    {
        EnsureWorkspace();

        _signatureHelpCts?.Cancel();
        _signatureHelpCts = new CancellationTokenSource();
        var ct = _signatureHelpCts.Token;

        var textUpToCaret = FormulaEditor.Document.GetText(0, FormulaEditor.CaretOffset);
        var fullText = FormulaEditor.Text;

        SignatureHelpModel? model;

        try
        {
            model = await _signatureHelpProvider!.GetSignatureHelpAsync(
                textUpToCaret, fullText, Metadata, ct);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (ct.IsCancellationRequested)
        {
            return;
        }

        if (model == null)
        {
            DismissSignatureHelp();
            return;
        }

        _signatureHelpPopup ??= new SignatureHelpPopup(FormulaEditor);
        _signatureHelpPopup.Update(model);
    }

    private void DismissSignatureHelp()
    {
        _signatureHelpCts?.Cancel();
        _signatureHelpPopup?.Hide();
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

    private void OnEditorPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (Keyboard.Modifiers != ModifierKeys.Control)
        {
            return;
        }

        e.Handled = true;

        const double minFontSize = 8;
        const double maxFontSize = 30;
        var delta = e.Delta > 0 ? 1.0 : -1.0;
        var newSize = Math.Clamp(FormulaEditor.FontSize + delta, minFontSize, maxFontSize);

        FormulaEditor.FontSize = newSize;
        _settings.FontSize = newSize;
        _settings.Save();
    }

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        FormulaApplied?.Invoke(this, FormulaText);
        Hide();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => Hide();

    private void ToggleWordWrap()
    {
        FormulaEditor.WordWrap = !FormulaEditor.WordWrap;
        _settings.WordWrap = FormulaEditor.WordWrap;
        _settings.Save();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        // Don't actually close, just hide - we reuse the window
        e.Cancel = true;
        DismissSignatureHelp();
        Hide();
    }

    internal void ApplySettings(EditorSettings settings)
    {
        _settings.CopyFrom(settings);
        FormulaEditor.WordWrap = settings.WordWrap;
        FormulaEditor.Options.IndentationSize = settings.IndentSize;
    }

    internal void DisposeWorkspace()
    {
        _completionCts?.Cancel();
        _completionCts?.Dispose();
        _completionCts = null;
        _signatureHelpCts?.Cancel();
        _signatureHelpCts?.Dispose();
        _signatureHelpCts = null;
        DismissSignatureHelp();
        _workspaceManager?.Dispose();
        _workspaceManager = null;
        _completionProvider = null;
        _signatureHelpProvider = null;
    }
}
