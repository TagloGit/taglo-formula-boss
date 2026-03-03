using System.Text.RegularExpressions;

using FormulaBoss.Parsing;

namespace FormulaBoss.UI;

/// <summary>
///     The inferred type of the expression to the left of the caret dot.
///     Used only for routing completions: Row contexts get column completions,
///     everything else falls through to Roslyn.
/// </summary>
public enum DslType
{
    /// <summary>A row parameter inside a .Rows() lambda — show column completions.</summary>
    Row,

    /// <summary>Could not determine — delegate to Roslyn.</summary>
    Unknown,

    /// <summary>Not after a dot — show top-level items (keywords, variables, tables).</summary>
    TopLevel
}

/// <summary>
///     Result of context resolution for intellisense.
/// </summary>
public record CompletionContext(
    DslType Type,
    string? PartialWord,
    bool InsideDsl = true,
    bool IsBracketContext = false,
    string? TableName = null);

/// <summary>
///     Resolves the DSL type context at the caret position using a token-based backward walk.
///     Its only job is to detect Row contexts for column completions and whether the caret
///     is inside a backtick region. All other completions are handled by Roslyn.
/// </summary>
public static class ContextResolver
{
    private static readonly Regex CellRefPattern = new(
        @"^[A-Za-z]{1,3}\d{1,7}$", RegexOptions.Compiled);

    /// <summary>
    ///     Resolves the type context at the caret position.
    /// </summary>
    public static CompletionContext Resolve(string textUpToCaret, WorkbookMetadata? metadata)
    {
        if (string.IsNullOrEmpty(textUpToCaret))
        {
            return new CompletionContext(DslType.TopLevel, null);
        }

        // Determine if the caret is inside a backtick DSL region
        var insideDsl = IsInsideBackticks(textUpToCaret);

        // If inside backticks, only lex the DSL portion (from last unmatched backtick)
        var textToLex = textUpToCaret;
        if (insideDsl)
        {
            var lastBacktick = textUpToCaret.LastIndexOf('`');
            if (lastBacktick >= 0)
            {
                textToLex = textUpToCaret[(lastBacktick + 1)..];
            }
        }

        List<Token> tokens;
        try
        {
            tokens = new Lexer(textToLex).ScanTokens();
        }
        catch
        {
            return new CompletionContext(DslType.Unknown, null, insideDsl);
        }

        // Remove EOF token
        if (tokens.Count > 0 && tokens[^1].Type == TokenType.Eof)
        {
            tokens.RemoveAt(tokens.Count - 1);
        }

        if (tokens.Count == 0)
        {
            return new CompletionContext(DslType.TopLevel, null, insideDsl);
        }

        // Determine if we're after a dot (with optional partial word)
        string? partialWord = null;
        int dotIndex;

        var last = tokens[^1];
        if (last.Type == TokenType.Dot)
        {
            // Caret is right after a dot: "expr."
            dotIndex = tokens.Count - 1;
        }
        else if (last.Type == TokenType.Identifier && tokens.Count >= 2 && tokens[^2].Type == TokenType.Dot)
        {
            // Caret is after "expr.partialWo"
            partialWord = last.Lexeme;
            dotIndex = tokens.Count - 2;
        }
        else if (last.Type == TokenType.LeftBracket &&
                 tokens.Count >= 2 && tokens[^2].Type == TokenType.Identifier)
        {
            // Caret is right after "param[" — check if param is a row lambda parameter
            var paramName = tokens[^2].Lexeme;
            var (isRow, tableName) = IsRowContext(tokens, tokens.Count - 2, paramName);
            if (isRow)
            {
                return new CompletionContext(DslType.Row, null, insideDsl, true, tableName);
            }

            return new CompletionContext(DslType.TopLevel, null, insideDsl);
        }
        else if (last.Type == TokenType.Identifier &&
                 tokens.Count >= 3 && tokens[^2].Type == TokenType.LeftBracket &&
                 tokens[^3].Type == TokenType.Identifier)
        {
            // Caret is after "param[partialWo" — check if it's a row bracket context
            var paramName = tokens[^3].Lexeme;
            var (isRow, tableName) = IsRowContext(tokens, tokens.Count - 3, paramName);
            if (isRow)
            {
                return new CompletionContext(DslType.Row, last.Lexeme, insideDsl, true, tableName);
            }

            return new CompletionContext(DslType.TopLevel, GetTrailingWord(tokens), insideDsl);
        }
        else
        {
            // Not after a dot — top level context
            return new CompletionContext(DslType.TopLevel, GetTrailingWord(tokens), insideDsl);
        }

        // Walk the chain backward from the dot to find the expression root
        var (chainType, chainTableName) = ResolveChainType(tokens, dotIndex, metadata);
        return new CompletionContext(chainType, partialWord, insideDsl, TableName: chainTableName);
    }

    /// <summary>
    ///     Determines if the caret position is inside an unmatched backtick (i.e., inside a DSL expression).
    /// </summary>
    private static bool IsInsideBackticks(string textUpToCaret)
    {
        var count = 0;
        foreach (var c in textUpToCaret)
        {
            if (c == '`')
            {
                count++;
            }
        }

        // Odd number of backticks means we're inside an unclosed one
        return count % 2 == 1;
    }

    private static (DslType Type, string? TableName) ResolveChainType(
        List<Token> tokens, int dotIndex, WorkbookMetadata? metadata)
    {
        // Collect the chain: walk left from dotIndex collecting Identifier, Dot, and method calls
        // Chain is built right-to-left, then reversed
        var chain = new List<string>();
        var pos = dotIndex - 1;

        while (pos >= 0)
        {
            // Skip right paren + balanced parens (method call arguments)
            if (tokens[pos].Type == TokenType.RightParen)
            {
                pos = SkipBalancedParens(tokens, pos);
                if (pos < 0)
                {
                    break;
                }

                // After skipping parens, expect an identifier (method name)
                if (tokens[pos].Type == TokenType.Identifier)
                {
                    chain.Add(tokens[pos].Lexeme);
                    pos--;
                }
                else
                {
                    break;
                }
            }
            else if (tokens[pos].Type == TokenType.Identifier)
            {
                chain.Add(tokens[pos].Lexeme);
                pos--;
            }
            else
            {
                break;
            }

            // Expect a dot before the next segment (or we've reached the root)
            if (pos >= 0 && tokens[pos].Type == TokenType.Dot)
            {
                pos--;
            }
            else
            {
                break;
            }
        }

        // Check if chain root starts with a range reference (e.g., A1:B10)
        var hasRangePrefix = false;
        if (pos >= 0 && tokens[pos].Type == TokenType.Colon)
        {
            if (pos >= 1 && tokens[pos - 1].Type == TokenType.Identifier)
            {
                var startRef = tokens[pos - 1].Lexeme;
                if (chain.Count > 0 && IsCellReference(startRef) && IsCellReference(chain[^1]))
                {
                    hasRangePrefix = true;
                    chain.RemoveAt(chain.Count - 1);
                }
            }
        }

        chain.Reverse();

        if (chain.Count == 0)
        {
            return (DslType.Unknown, null);
        }

        // Check if any segment in the chain is ".Rows" — that makes this a Row context
        var root = chain[0];
        var startIdx = 0;

        if (hasRangePrefix || IsRangeIdentifier(root, metadata))
        {
            // Root is a known range/table — skip it when scanning for .Rows
            if (!hasRangePrefix)
            {
                startIdx = 1;
            }
        }
        else
        {
            // Check if root is a lambda parameter in a Row context
            var (isRow, tableName) = IsRowContext(tokens, dotIndex, root);
            if (isRow)
            {
                // If chain is just the param name (e.g. "r."), show column completions
                if (chain.Count == 1)
                {
                    return (DslType.Row, tableName);
                }

                // After a member access on a row param (e.g. r.Price.) — Roslyn handles it
                return (DslType.Unknown, null);
            }
        }

        // Scan the chain for the last .Rows accessor — if the chain ends with .Rows,
        // the caret is in a Row context
        for (var i = chain.Count - 1; i >= startIdx; i--)
        {
            if (chain[i].Equals("Rows", StringComparison.OrdinalIgnoreCase))
            {
                // .Rows is the last meaningful segment — Row context
                if (i == chain.Count - 1)
                {
                    // Find root table name for column completions
                    var tableName = hasRangePrefix ? null : root;
                    return (DslType.Row, tableName);
                }

                // Something after .Rows (e.g. .Rows.Where(...).) — still could be Row
                // if a later lambda parameter is in scope, but for the chain walk
                // the result is Unknown (Roslyn handles method completions)
                break;
            }
        }

        return (DslType.Unknown, null);
    }

    /// <summary>
    ///     Determines whether a lambda parameter is in a Row context
    ///     by scanning backward for the enclosing .Rows accessor in the chain.
    /// </summary>
    private static (bool IsRow, string? TableName) IsRowContext(
        List<Token> tokens, int dotIndex, string paramName)
    {
        // Scan backward from dotIndex looking for an unmatched `(` that belongs to a
        // method call, then check if paramName is declared as a lambda parameter there.
        var depth = 0;

        for (var i = dotIndex - 1; i >= 0; i--)
        {
            var tok = tokens[i];

            if (tok.Type == TokenType.RightParen)
            {
                depth++;
            }
            else if (tok.Type == TokenType.LeftParen)
            {
                if (depth > 0)
                {
                    depth--;
                }
                else
                {
                    // Found unmatched open paren — check if paramName is a lambda param here
                    if (!IsLambdaParamInScope(tokens, i, dotIndex, paramName))
                    {
                        continue;
                    }

                    // The method name is before this paren
                    if (i >= 2 && tokens[i - 1].Type == TokenType.Identifier &&
                        tokens[i - 2].Type == TokenType.Dot)
                    {
                        var methodName = tokens[i - 1].Lexeme;

                        if (methodName.Equals("Rows", StringComparison.OrdinalIgnoreCase))
                        {
                            return (true, FindChainRootTable(tokens, i - 2));
                        }

                        // Any other method — continue walking backward to find .Rows
                        return FindRowsInChain(tokens, i - 2);
                    }
                }
            }
        }

        return (false, null);
    }

    /// <summary>
    ///     Checks whether paramName appears as a lambda parameter between the opening paren
    ///     and an arrow token (=>) within the given range.
    /// </summary>
    private static bool IsLambdaParamInScope(List<Token> tokens, int leftParenIndex, int dotIndex,
        string paramName)
    {
        for (var j = leftParenIndex + 1; j < dotIndex; j++)
        {
            if (tokens[j].Type == TokenType.Arrow)
            {
                for (var k = leftParenIndex + 1; k < j; k++)
                {
                    if (tokens[k].Type == TokenType.Identifier &&
                        tokens[k].Lexeme.Equals(paramName, StringComparison.Ordinal))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        return false;
    }

    /// <summary>
    ///     Walks backward through a method chain looking for .Rows to determine
    ///     if a lambda parameter is in a Row context.
    /// </summary>
    private static (bool IsRow, string? TableName) FindRowsInChain(
        List<Token> tokens, int dotPos)
    {
        var pos = dotPos - 1;

        while (pos >= 0)
        {
            if (tokens[pos].Type == TokenType.RightParen)
            {
                pos = SkipBalancedParens(tokens, pos);
                if (pos < 0)
                {
                    break;
                }

                if (tokens[pos].Type == TokenType.Identifier)
                {
                    if (tokens[pos].Lexeme.Equals("Rows", StringComparison.OrdinalIgnoreCase))
                    {
                        return (true, FindChainRootTable(tokens, pos));
                    }

                    pos--;
                }
                else
                {
                    break;
                }
            }
            else if (tokens[pos].Type == TokenType.Identifier)
            {
                if (tokens[pos].Lexeme.Equals("Rows", StringComparison.OrdinalIgnoreCase))
                {
                    return (true, FindChainRootTable(tokens, pos));
                }

                pos--;
            }
            else if (tokens[pos].Type == TokenType.Dot)
            {
                pos--;
            }
            else
            {
                break;
            }
        }

        return (false, null);
    }

    /// <summary>
    ///     From a position inside a chain (at .Rows),
    ///     walks backward past the dot to find the root identifier (table name).
    /// </summary>
    private static string? FindChainRootTable(List<Token> tokens, int accessorPos)
    {
        var pos = accessorPos - 1;

        while (pos >= 0)
        {
            if (tokens[pos].Type == TokenType.Dot)
            {
                pos--;
                if (pos >= 0 && tokens[pos].Type == TokenType.RightParen)
                {
                    pos = SkipBalancedParens(tokens, pos);
                    if (pos < 0)
                    {
                        break;
                    }

                    if (tokens[pos].Type == TokenType.Identifier)
                    {
                        pos--;
                        continue;
                    }

                    break;
                }

                if (pos >= 0 && tokens[pos].Type == TokenType.Identifier)
                {
                    var candidate = tokens[pos].Lexeme;
                    if (pos == 0 || tokens[pos - 1].Type != TokenType.Dot)
                    {
                        return candidate;
                    }

                    pos--;
                    continue;
                }
            }

            break;
        }

        return null;
    }

    /// <summary>
    ///     Skips backward over balanced parentheses, returning the index of the matching left paren - 1.
    ///     Returns -1 if unbalanced.
    /// </summary>
    private static int SkipBalancedParens(List<Token> tokens, int rightParenIndex)
    {
        var depth = 1;
        var pos = rightParenIndex - 1;
        while (pos >= 0 && depth > 0)
        {
            if (tokens[pos].Type == TokenType.RightParen)
            {
                depth++;
            }
            else if (tokens[pos].Type == TokenType.LeftParen)
            {
                depth--;
            }

            if (depth > 0)
            {
                pos--;
            }
        }

        return pos >= 0 ? pos - 1 : -1;
    }

    private static bool IsRangeIdentifier(string name, WorkbookMetadata? metadata)
    {
        if (IsCellReference(name))
        {
            return true;
        }

        if (metadata == null)
        {
            return false;
        }

        foreach (var t in metadata.TableNames)
        {
            if (t.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        foreach (var n in metadata.NamedRanges)
        {
            if (n.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsCellReference(string text)
    {
        var stripped = text.Replace("$", "");
        return CellRefPattern.IsMatch(stripped);
    }

    private static string? GetTrailingWord(List<Token> tokens)
    {
        if (tokens.Count > 0 && tokens[^1].Type == TokenType.Identifier)
        {
            return tokens[^1].Lexeme;
        }

        return null;
    }
}
