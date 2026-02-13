using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;

namespace FormulaBoss.UI;

/// <summary>
///     Static helpers for code-editing behaviors (brace completion, indentation).
///     Operates on AvalonEdit TextEditor/TextDocument to keep event handlers thin.
/// </summary>
internal static class EditorBehaviors
{
    private static readonly Dictionary<char, char> BracePairs = new()
    {
        { '(', ')' }, { '[', ']' }, { '{', '}' }, { '"', '"' }, { '`', '`' }
    };

    /// <summary>
    ///     If the typed character is a closing brace/quote that already exists at the caret,
    ///     skip over it instead of inserting a duplicate. Returns true if handled.
    /// </summary>
    public static bool TrySkipClosingChar(TextEditor editor, char ch)
    {
        if (ch is not (')' or ']' or '}' or '"' or '`'))
        {
            return false;
        }

        var offset = editor.CaretOffset;
        var doc = editor.Document;

        if (offset >= doc.TextLength || doc.GetCharAt(offset) != ch)
        {
            return false;
        }

        editor.CaretOffset = offset + 1;
        return true;
    }

    /// <summary>
    ///     After an opening brace/quote is inserted, auto-insert the matching closer
    ///     and position the caret between the pair.
    /// </summary>
    public static void AutoInsertClosingChar(TextEditor editor, char ch)
    {
        if (!BracePairs.TryGetValue(ch, out var closingChar))
        {
            return;
        }

        var offset = editor.CaretOffset;
        editor.Document.Insert(offset, closingChar.ToString());
        editor.CaretOffset = offset;
    }

    /// <summary>
    ///     When '{' is typed on a line with only whitespace before it,
    ///     de-indent to match the previous non-empty line's indent level.
    /// </summary>
    public static void DeIndentOpenBrace(TextEditor editor)
    {
        var doc = editor.Document;
        var line = doc.GetLineByOffset(editor.CaretOffset);
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
        editor.CaretOffset = line.Offset + prevIndent.Length + 1;
    }

    /// <summary>
    ///     When Enter is pressed between {}, expand into three lines with indentation.
    ///     Returns true if handled.
    /// </summary>
    public static bool TryExpandBraceBlock(TextEditor editor)
    {
        var offset = editor.CaretOffset;
        var doc = editor.Document;

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
        var innerIndent = baseIndent + new string(' ', editor.Options.IndentationSize);

        var insertion = "\r\n" + innerIndent + "\r\n" + baseIndent;
        doc.Insert(offset, insertion);
        editor.CaretOffset = offset + "\r\n".Length + innerIndent.Length;
        return true;
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
