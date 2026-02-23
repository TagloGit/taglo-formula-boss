using System.Diagnostics;
using System.Windows.Input;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;

namespace FormulaBoss.UI;

/// <summary>
///     Manages code-editing behaviors (brace completion, indentation, key handling)
///     for an AvalonEdit TextEditor instance. Subscribes to editor events internally
///     and exposes callbacks for actions that require window-level coordination.
/// </summary>
internal class EditorBehaviorHandler
{
    private static readonly Dictionary<char, char> BracePairs = new()
    {
        { '(', ')' },
        { '[', ']' },
        { '{', '}' },
        { '"', '"' },
        { '`', '`' }
    };

    private readonly TextEditor _editor;

    public EditorBehaviorHandler(TextEditor editor)
    {
        _editor = editor;

        _editor.TextArea.TextEntering += OnTextEntering;
        _editor.TextArea.TextEntered += OnTextEntered;
        _editor.PreviewKeyDown += OnPreviewKeyDown;
    }

    /// <summary>
    ///     Invoked when the editor wants to trigger intellisense completion.
    ///     The char parameter is the trigger character ('.', '[', or a letter).
    /// </summary>
    public Action<char>? CompletionRequested { get; set; }

    /// <summary>
    ///     Invoked when a non-identifier character is typed and the completion
    ///     window should be closed. Passes the typed character.
    /// </summary>
    public Action<char>? CompletionCloseRequested { get; set; }

    /// <summary>
    ///     Invoked when Ctrl+Space is pressed and completion should be force-shown.
    /// </summary>
    public Action? ForceCompletionRequested { get; set; }

    /// <summary>
    ///     Invoked when Ctrl+Enter is pressed to apply the formula.
    /// </summary>
    public Action<string>? FormulaApplyRequested { get; set; }

    /// <summary>
    ///     Callback to check whether the completion window has any visible items.
    ///     Returns true if the completion list is empty and should be closed.
    /// </summary>
    public Func<bool>? IsCompletionListEmpty { get; set; }

    /// <summary>
    ///     Whether the current completion context is bracket-based (e.g. [col name]),
    ///     which allows spaces in completion filtering.
    /// </summary>
    public bool IsBracketContext { get; set; }

    private void OnTextEntering(object sender, TextCompositionEventArgs e)
    {
        try
        {
            if (e.Text.Length != 1)
            {
                return;
            }

            var ch = e.Text[0];

            if (!char.IsLetterOrDigit(ch) && !(ch == ' ' && IsBracketContext))
            {
                CompletionCloseRequested?.Invoke(ch);
            }

            if (!e.Handled && TrySkipClosingChar(ch))
            {
                e.Handled = true;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"OnTextEntering error: {ex.Message}");
        }
    }

    private void OnTextEntered(object sender, TextCompositionEventArgs e)
    {
        try
        {
            if (e.Text.Length != 1)
            {
                return;
            }

            var ch = e.Text[0];
            if (ch == '{')
            {
                DeIndentOpenBrace();
            }

            AutoInsertClosingChar(ch);

            if (ch == '.' || ch == '[' || char.IsLetter(ch))
            {
                CompletionRequested?.Invoke(ch);
            }

            if (IsCompletionListEmpty?.Invoke() == true)
            {
                CompletionCloseRequested?.Invoke(ch);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"OnTextEntered error: {ex.Message}");
        }
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        try
        {
            if (e is { Key: Key.Space, KeyboardDevice.Modifiers: ModifierKeys.Control })
            {
                e.Handled = true;
                ForceCompletionRequested?.Invoke();
                return;
            }

            if (e.Key == Key.Enter && e.KeyboardDevice.Modifiers == ModifierKeys.None
                                   && (TryExpandBraceBlock() || TryExpandBeforeClosingParen()))
            {
                e.Handled = true;
                return;
            }

            if (e is { Key: Key.Enter, KeyboardDevice.Modifiers: ModifierKeys.Control })
            {
                e.Handled = true;
                FormulaApplyRequested?.Invoke(_editor.Text);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"OnPreviewKeyDown error: {ex.Message}");
        }
    }

    /// <summary>
    ///     If the typed character is a closing brace/quote that already exists at the caret,
    ///     skip over it instead of inserting a duplicate. Returns true if handled.
    /// </summary>
    internal bool TrySkipClosingChar(char ch)
    {
        if (ch is not (')' or ']' or '}' or '"' or '`'))
        {
            return false;
        }

        var offset = _editor.CaretOffset;
        var doc = _editor.Document;

        if (offset >= doc.TextLength || doc.GetCharAt(offset) != ch)
        {
            return false;
        }

        _editor.CaretOffset = offset + 1;
        return true;
    }

    /// <summary>
    ///     After an opening brace/quote is inserted, auto-insert the matching closer
    ///     and position the caret between the pair.
    /// </summary>
    internal void AutoInsertClosingChar(char ch)
    {
        if (!BracePairs.TryGetValue(ch, out var closingChar))
        {
            return;
        }

        var offset = _editor.CaretOffset;
        _editor.Document.Insert(offset, closingChar.ToString());
        _editor.CaretOffset = offset;
    }

    /// <summary>
    ///     When '{' is typed on a line with only whitespace before it,
    ///     de-indent to match the previous non-empty line's indent level.
    /// </summary>
    internal void DeIndentOpenBrace()
    {
        var doc = _editor.Document;
        var line = doc.GetLineByOffset(_editor.CaretOffset);
        var lineText = doc.GetText(line);
        var trimmed = lineText.TrimStart();

        if (trimmed is not ("{}" or "{"))
        {
            return;
        }

        var prevIndent = GetPreviousLineIndent(doc, line.LineNumber);
        var currentIndent = lineText[..^trimmed.Length];
        if (currentIndent == prevIndent)
        {
            return;
        }

        doc.Replace(line.Offset, line.Length, prevIndent + trimmed);
        _editor.CaretOffset = line.Offset + prevIndent.Length + 1;
    }

    /// <summary>
    ///     When Enter is pressed between {}, expand into three lines with indentation.
    ///     Returns true if handled.
    /// </summary>
    internal bool TryExpandBraceBlock()
    {
        var offset = _editor.CaretOffset;
        var doc = _editor.Document;

        if (offset <= 0 || offset >= doc.TextLength)
        {
            return false;
        }

        if (doc.GetCharAt(offset - 1) != '{' || doc.GetCharAt(offset) != '}')
        {
            return false;
        }

        var line = doc.GetLineByOffset(offset);
        var lineText = doc.GetText(line);
        var baseIndent = lineText[..^lineText.TrimStart().Length];
        var innerIndent = baseIndent + new string(' ', _editor.Options.IndentationSize);

        var insertion = "\r\n" + innerIndent + "\r\n" + baseIndent;
        doc.Insert(offset, insertion);
        _editor.CaretOffset = offset + "\r\n".Length + innerIndent.Length;
        return true;
    }

    /// <summary>
    ///     When Enter is pressed immediately before a closing paren/bracket,
    ///     move the closer to the next line with matching indentation and place
    ///     the caret on a new indented line between. Returns true if handled.
    /// </summary>
    internal bool TryExpandBeforeClosingParen()
    {
        var offset = _editor.CaretOffset;
        var doc = _editor.Document;

        if (offset >= doc.TextLength)
        {
            return false;
        }

        var nextChar = doc.GetCharAt(offset);
        if (nextChar is not (')' or ']'))
        {
            return false;
        }

        var openChar = nextChar == ')' ? '(' : '[';
        var openerLine = FindMatchingOpen(doc, offset, openChar, nextChar);
        var openerLineText = doc.GetText(doc.GetLineByNumber(openerLine));
        var baseIndent = openerLineText[..^openerLineText.TrimStart().Length];
        var indent = baseIndent + new string(' ', _editor.Options.IndentationSize);

        var insertion = "\r\n" + indent;
        doc.Insert(offset, insertion);
        _editor.CaretOffset = offset + insertion.Length;
        return true;
    }

    private static int FindMatchingOpen(TextDocument doc, int closeOffset, char open, char close)
    {
        var depth = 1;
        for (var i = closeOffset - 1; i >= 0; i--)
        {
            var ch = doc.GetCharAt(i);
            if (ch == close)
            {
                depth++;
            }
            else if (ch == open)
            {
                depth--;
            }

            if (depth == 0)
            {
                return doc.GetLineByOffset(i).LineNumber;
            }
        }

        return 1;
    }

    private static string GetPreviousLineIndent(TextDocument doc, int lineNumber)
    {
        for (var i = lineNumber - 1; i >= 1; i--)
        {
            var prevLine = doc.GetLineByNumber(i);
            var prevText = doc.GetText(prevLine);
            if (prevText.Trim().Length > 0)
            {
                return prevText[..^prevText.TrimStart().Length];
            }
        }

        return "";
    }
}
