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

    /// <summary>
    ///     Returns true when the completion window is open, so Tab can be
    ///     passed through for completion selection instead of indenting.
    /// </summary>
    public Func<bool>? IsCompletionWindowOpen { get; set; }

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

            if (!e.Handled && TrySurroundSelection(ch))
            {
                e.Handled = true;
            }
            else if (!e.Handled && TrySkipClosingChar(ch))
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
                return;
            }

            if (e.Key == Key.Home && e.KeyboardDevice.Modifiers is ModifierKeys.None or ModifierKeys.Shift)
            {
                var extend = e.KeyboardDevice.Modifiers == ModifierKeys.Shift;
                SmartHome(extend);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Tab && e.KeyboardDevice.Modifiers is ModifierKeys.None or ModifierKeys.Shift
                                 && IsCompletionWindowOpen?.Invoke() != true)
            {
                var shift = e.KeyboardDevice.Modifiers == ModifierKeys.Shift;
                if (shift)
                {
                    Dedent();
                }
                else
                {
                    Indent();
                }

                e.Handled = true;
                return;
            }

            if (e.Key == Key.Back && e.KeyboardDevice.Modifiers == ModifierKeys.None)
            {
                if (TryDeletePairedChars() || TrySmartBackspace())
                {
                    e.Handled = true;
                }
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
    ///     and position the caret between the pair. For quotes (" and `), suppresses
    ///     auto-close when the next character is an identifier char, digit, underscore,
    ///     or another quote — only auto-closes before whitespace, closing brackets,
    ///     commas, semicolons, or end-of-line.
    /// </summary>
    internal void AutoInsertClosingChar(char ch)
    {
        if (!BracePairs.TryGetValue(ch, out var closingChar))
        {
            return;
        }

        if (ch is '"' or '`' && !ShouldAutoCloseQuote(ch))
        {
            return;
        }

        var offset = _editor.CaretOffset;
        _editor.Document.Insert(offset, closingChar.ToString());
        _editor.CaretOffset = offset;
    }

    /// <summary>
    ///     When text is selected and the user types an opener, wrap the selection
    ///     with the pair instead of replacing it. Returns true if handled.
    /// </summary>
    internal bool TrySurroundSelection(char ch)
    {
        if (_editor.SelectionLength == 0)
        {
            return false;
        }

        if (!BracePairs.TryGetValue(ch, out var closingChar))
        {
            return false;
        }

        var start = _editor.SelectionStart;
        var selectedText = _editor.SelectedText;
        var replacement = ch + selectedText + closingChar;

        _editor.Document.Replace(start, selectedText.Length, replacement);
        // Select the wrapped content (excluding the pair chars)
        _editor.Select(start + 1, selectedText.Length);
        return true;
    }

    /// <summary>
    ///     Returns true if a quote should be auto-closed. Suppresses auto-close when
    ///     there is an unmatched quote of the same type before the caret on the current
    ///     line (i.e. we're inside a string). When not inside a string, only auto-closes
    ///     before whitespace, closing brackets, commas, semicolons, or end-of-line.
    /// </summary>
    private bool ShouldAutoCloseQuote(char quoteChar)
    {
        var offset = _editor.CaretOffset;
        var doc = _editor.Document;
        var line = doc.GetLineByOffset(offset);

        // Count occurrences of this quote char before caret on the current line.
        // The char that was just inserted is at offset-1, so look before that.
        var lineTextBeforeInserted = doc.GetText(line.Offset, offset - 1 - line.Offset);
        var count = 0;
        foreach (var c in lineTextBeforeInserted)
        {
            if (c == quoteChar)
            {
                count++;
            }
        }

        // Odd count means we're inside an open quote — don't auto-close
        if (count % 2 != 0)
        {
            return false;
        }

        // Even count (or zero) — apply forward-looking heuristic
        if (offset >= doc.TextLength)
        {
            return true;
        }

        var next = doc.GetCharAt(offset);
        return next is ' ' or '\t' or '\r' or '\n' or ')' or ']' or '}' or ',' or ';';
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

    /// <summary>
    ///     Smart Home: first press jumps to first non-whitespace character on the line;
    ///     second press (or if already there) jumps to column 0. Toggles between the two.
    ///     When extend is true, the selection is extended rather than moving the caret.
    /// </summary>
    internal void SmartHome(bool extend)
    {
        var doc = _editor.Document;
        var line = doc.GetLineByOffset(_editor.CaretOffset);
        var lineText = doc.GetText(line);
        var firstNonWhitespace = line.Offset + (lineText.Length - lineText.TrimStart().Length);

        var target = _editor.CaretOffset == firstNonWhitespace ? line.Offset : firstNonWhitespace;

        if (extend)
        {
            var anchor = _editor.TextArea.Selection.IsEmpty
                ? _editor.CaretOffset
                : _editor.SelectionStart == _editor.CaretOffset
                    ? _editor.SelectionStart + _editor.SelectionLength
                    : _editor.SelectionStart;
            var selStart = Math.Min(anchor, target);
            var selLength = Math.Abs(anchor - target);
            _editor.Select(selStart, selLength);
            // Place caret at the target end of the selection
            _editor.CaretOffset = target;
        }
        else
        {
            _editor.CaretOffset = target;
        }
    }

    /// <summary>
    ///     Smart Backspace: when the caret is in leading whitespace, remove back to
    ///     the previous indent stop (nearest multiple of indent size) instead of one char.
    ///     Returns true if handled.
    /// </summary>
    internal bool TrySmartBackspace()
    {
        var doc = _editor.Document;
        var offset = _editor.CaretOffset;
        if (offset == 0)
        {
            return false;
        }

        var line = doc.GetLineByOffset(offset);
        var colInLine = offset - line.Offset;
        var lineText = doc.GetText(line);

        // Only act in leading whitespace (all chars before caret are spaces)
        if (colInLine == 0 || lineText[..colInLine].TrimStart().Length > 0)
        {
            return false;
        }

        var indentSize = _editor.Options.IndentationSize;
        // Calculate previous indent stop
        var targetCol = colInLine <= indentSize
            ? 0
            : (colInLine - 1) / indentSize * indentSize;

        var removeCount = colInLine - targetCol;
        doc.Remove(line.Offset + targetCol, removeCount);
        _editor.CaretOffset = line.Offset + targetCol;
        return true;
    }

    /// <summary>
    ///     Paired Delete: when Backspace is pressed between an empty auto-inserted pair
    ///     (e.g. (), [], {}, "", ``), delete both characters. Returns true if handled.
    /// </summary>
    internal bool TryDeletePairedChars()
    {
        var offset = _editor.CaretOffset;
        var doc = _editor.Document;

        if (offset <= 0 || offset >= doc.TextLength)
        {
            return false;
        }

        var before = doc.GetCharAt(offset - 1);
        var after = doc.GetCharAt(offset);

        if (!BracePairs.TryGetValue(before, out var expected) || after != expected)
        {
            return false;
        }

        doc.Remove(offset - 1, 2);
        _editor.CaretOffset = offset - 1;
        return true;
    }

    /// <summary>
    ///     Tab with no multiline selection: insert spaces to the next tab stop.
    ///     Tab with multiline selection: indent all selected lines by one indent unit.
    ///     Selection is preserved after indent.
    /// </summary>
    internal void Indent()
    {
        var doc = _editor.Document;
        var indentSize = _editor.Options.IndentationSize;
        var indent = new string(' ', indentSize);
        var selStart = _editor.SelectionStart;
        var selLength = _editor.SelectionLength;

        var startLine = doc.GetLineByOffset(selStart);
        var endLine = doc.GetLineByOffset(selStart + selLength);

        if (selLength == 0 || startLine.LineNumber == endLine.LineNumber)
        {
            // No selection or single-line selection: insert spaces to next tab stop
            var offset = _editor.CaretOffset;
            var line = doc.GetLineByOffset(offset);
            var col = offset - line.Offset;
            var spacesToInsert = indentSize - col % indentSize;
            doc.Insert(offset, new string(' ', spacesToInsert));
            _editor.CaretOffset = offset + spacesToInsert;
            return;
        }

        // Multiline selection: indent all lines
        doc.BeginUpdate();
        var addedTotal = 0;
        for (var i = startLine.LineNumber; i <= endLine.LineNumber; i++)
        {
            var line = doc.GetLineByNumber(i);
            doc.Insert(line.Offset, indent);
            addedTotal += indentSize;
        }

        doc.EndUpdate();

        // Preserve selection spanning the same lines
        var newSelStart = startLine.Offset;
        var lastLine = doc.GetLineByNumber(endLine.LineNumber);
        _editor.Select(newSelStart, lastLine.EndOffset - newSelStart);
    }

    /// <summary>
    ///     Shift+Tab: remove one indent level from the current line or all selected lines.
    ///     If a line has fewer spaces than one indent unit, remove all leading whitespace.
    ///     Selection is preserved after dedent.
    /// </summary>
    internal void Dedent()
    {
        var doc = _editor.Document;
        var indentSize = _editor.Options.IndentationSize;
        var selStart = _editor.SelectionStart;
        var selLength = _editor.SelectionLength;

        var startLine = doc.GetLineByOffset(selStart);
        var endLine = doc.GetLineByOffset(selStart + selLength);

        doc.BeginUpdate();
        for (var i = startLine.LineNumber; i <= endLine.LineNumber; i++)
        {
            var line = doc.GetLineByNumber(i);
            var lineText = doc.GetText(line);
            var leadingSpaces = 0;
            foreach (var c in lineText)
            {
                if (c == ' ')
                {
                    leadingSpaces++;
                }
                else
                {
                    break;
                }
            }

            var toRemove = Math.Min(leadingSpaces, indentSize);
            if (toRemove > 0)
            {
                doc.Remove(line.Offset, toRemove);
            }
        }

        doc.EndUpdate();

        // Preserve selection spanning the same lines
        if (selLength > 0 && startLine.LineNumber != endLine.LineNumber)
        {
            var newStart = doc.GetLineByNumber(startLine.LineNumber).Offset;
            var lastLine = doc.GetLineByNumber(endLine.LineNumber);
            _editor.Select(newStart, lastLine.EndOffset - newStart);
        }
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
