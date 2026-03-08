using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
///     Splits LET formula arguments by comma, correctly handling nested brackets,
///     backtick regions (treated as opaque), and string literals.
/// </summary>
internal static class LetArgumentSplitter
{
    /// <summary>
    ///     Splits LET arguments from a body string (content between the outer LET parentheses).
    ///     Requires the body to be already extracted (no outer parens).
    /// </summary>
    public static List<string> Split(string body)
    {
        return SplitCore(body, 0, requireClosingParen: false);
    }

    /// <summary>
    ///     Splits LET arguments starting at <paramref name="startPos"/> in <paramref name="text"/>,
    ///     stopping at the matching closing parenthesis. Tolerant of incomplete input
    ///     (missing closing paren returns whatever arguments were found).
    /// </summary>
    public static List<string> SplitTolerant(string text, int startPos)
    {
        return SplitCore(text, startPos, requireClosingParen: false);
    }

    /// <summary>
    ///     Finds the matching closing parenthesis for the opening parenthesis at
    ///     <paramref name="openParenIndex"/>, correctly skipping nested brackets,
    ///     backtick regions, and string literals.
    /// </summary>
    /// <returns>Index of the matching closing parenthesis, or -1 if not found.</returns>
    public static int FindMatchingCloseParen(string text, int openParenIndex)
    {
        var depth = 0;
        var inBacktick = false;
        var inString = false;
        var stringChar = '\0';

        for (var i = openParenIndex; i < text.Length; i++)
        {
            var c = text[i];

            if (inBacktick)
            {
                if (c == '`')
                    inBacktick = false;
                continue;
            }

            if (inString)
            {
                if (c == stringChar)
                {
                    if (i + 1 < text.Length && text[i + 1] == stringChar)
                        i++; // Skip doubled quote
                    else
                        inString = false;
                }
                continue;
            }

            switch (c)
            {
                case '`':
                    inBacktick = true;
                    break;
                case '"':
                case '\'':
                    inString = true;
                    stringChar = c;
                    break;
                case '(':
                case '{':
                case '[':
                    depth++;
                    break;
                case ')':
                case '}':
                case ']':
                    depth--;
                    if (depth == 0 && c == ')')
                        return i;
                    break;
            }
        }

        return -1;
    }

    private static List<string> SplitCore(string text, int startPos, bool requireClosingParen)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var depth = 0;
        var inBacktick = false;
        var inString = false;
        var stringChar = '\0';

        for (var i = startPos; i < text.Length; i++)
        {
            var c = text[i];

            // Inside a backtick region: treat as opaque, skip everything
            if (inBacktick)
            {
                current.Append(c);
                if (c == '`')
                    inBacktick = false;
                continue;
            }

            // Inside a string literal: skip until closing quote
            if (inString)
            {
                current.Append(c);
                if (c == stringChar)
                {
                    if (i + 1 < text.Length && text[i + 1] == stringChar)
                    {
                        current.Append(text[i + 1]);
                        i++; // Skip doubled quote
                    }
                    else
                    {
                        inString = false;
                    }
                }
                continue;
            }

            switch (c)
            {
                case '`':
                    inBacktick = true;
                    current.Append(c);
                    break;
                case '"':
                case '\'':
                    inString = true;
                    stringChar = c;
                    current.Append(c);
                    break;
                case '(':
                case '{':
                case '[':
                    depth++;
                    current.Append(c);
                    break;
                case '}':
                case ']':
                    depth--;
                    current.Append(c);
                    break;
                case ')':
                    if (depth > 0)
                    {
                        depth--;
                        current.Append(c);
                    }
                    else
                    {
                        // Matching close paren for the outer LET( — stop here
                        if (current.Length > 0)
                            args.Add(current.ToString());
                        return args;
                    }
                    break;
                case ',' when depth == 0:
                    args.Add(current.ToString());
                    current.Clear();
                    break;
                default:
                    current.Append(c);
                    break;
            }
        }

        // Reached end of text without closing paren
        if (current.Length > 0)
            args.Add(current.ToString());

        return args;
    }
}
