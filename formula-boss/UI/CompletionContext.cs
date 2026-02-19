using System.Text.RegularExpressions;

using FormulaBoss.Parsing;

namespace FormulaBoss.UI;

/// <summary>
///     The inferred type of the expression to the left of the caret dot.
/// </summary>
public enum DslType
{
    /// <summary>An Excel range, table, or named range — needs an accessor next.</summary>
    Range,

    /// <summary>A pipeline (range + accessor + optional method chain) — can chain methods.</summary>
    Pipeline,

    /// <summary>A cell parameter inside a .cells() lambda.</summary>
    Cell,

    /// <summary>A row parameter inside a .rows() lambda.</summary>
    Row,

    /// <summary>A Cell.Interior sub-object.</summary>
    Interior,

    /// <summary>A Cell.Font sub-object.</summary>
    Font,

    /// <summary>A scalar result (after terminal aggregation or property access).</summary>
    Scalar,

    /// <summary>Could not determine — show fallback completions.</summary>
    Unknown,

    /// <summary>Not after a dot — show top-level items (keywords, variables, tables).</summary>
    TopLevel
}

/// <summary>
///     Result of context resolution for intellisense.
/// </summary>
public record CompletionContext(DslType Type, string? PartialWord, bool InsideDsl = true, bool IsBracketContext = false);

/// <summary>
///     Resolves the DSL type context at the caret position using a token-based backward walk.
///     Works on incomplete/invalid input since it uses only the lexer, not the parser.
/// </summary>
public static class ContextResolver
{
    private static readonly HashSet<string> Accessors = new(StringComparer.OrdinalIgnoreCase)
    {
        "cells", "values", "rows", "cols"
    };

    private static readonly HashSet<string> PipelineMethods = new(StringComparer.OrdinalIgnoreCase)
    {
        "where",
        "select",
        "map",
        "find",
        "some",
        "every",
        "orderBy",
        "orderByDesc",
        "take",
        "skip",
        "distinct",
        "groupBy",
        "reduce",
        "aggregate",
        "scan",
        "toArray"
    };

    private static readonly HashSet<string> TerminalMethods = new(StringComparer.OrdinalIgnoreCase)
    {
        "sum",
        "avg",
        "average",
        "min",
        "max",
        "count",
        "first",
        "firstOrDefault",
        "last",
        "lastOrDefault"
    };

    private static readonly HashSet<string> CellProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "value",
        "color",
        "rgb",
        "bold",
        "italic",
        "fontSize",
        "format",
        "formula",
        "row",
        "col",
        "address",
        "Value",
        "Row",
        "Column",
        "Address",
        "Formula",
        "NumberFormat"
    };

    private static readonly HashSet<string> CellContextMethods = new(StringComparer.OrdinalIgnoreCase) { "cells" };

    private static readonly HashSet<string> RowContextMethods = new(StringComparer.OrdinalIgnoreCase) { "rows" };

    // Methods whose lambda parameter inherits the pipeline element type
    private static readonly HashSet<string> PassthroughMethods = new(StringComparer.OrdinalIgnoreCase)
    {
        "where",
        "find",
        "some",
        "every",
        "select",
        "map",
        "orderBy",
        "orderByDesc",
        "scan"
    };

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
            var lambdaType = ResolveLambdaParamType(tokens, tokens.Count - 2, paramName);
            if (lambdaType == DslType.Row)
            {
                return new CompletionContext(DslType.Row, null, insideDsl, IsBracketContext: true);
            }

            return new CompletionContext(DslType.TopLevel, null, insideDsl);
        }
        else if (last.Type == TokenType.Identifier &&
                 tokens.Count >= 3 && tokens[^2].Type == TokenType.LeftBracket &&
                 tokens[^3].Type == TokenType.Identifier)
        {
            // Caret is after "param[partialWo" — check if it's a row bracket context
            var paramName = tokens[^3].Lexeme;
            var lambdaType = ResolveLambdaParamType(tokens, tokens.Count - 3, paramName);
            if (lambdaType == DslType.Row)
            {
                return new CompletionContext(DslType.Row, last.Lexeme, insideDsl, IsBracketContext: true);
            }

            return new CompletionContext(DslType.TopLevel, GetTrailingWord(tokens), insideDsl);
        }
        else
        {
            // Not after a dot — top level context
            return new CompletionContext(DslType.TopLevel, GetTrailingWord(tokens), insideDsl);
        }

        // Walk the chain backward from the dot to find the expression root
        var chainType = ResolveChainType(tokens, dotIndex, metadata);
        return new CompletionContext(chainType, partialWord, insideDsl);
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

    /// <summary>
    ///     Resolves the type of the expression chain ending at the given dot token index.
    /// </summary>
    private static DslType ResolveChainType(List<Token> tokens, int dotIndex, WorkbookMetadata? metadata)
    {
        // Collect the chain: walk left from dotIndex collecting Identifier, Dot, and method calls
        // Chain is built right-to-left, then reversed
        var chain = new List<ChainSegment>();
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
                    chain.Add(new ChainSegment(tokens[pos].Lexeme));
                    pos--;
                }
                else
                {
                    break;
                }
            }
            else if (tokens[pos].Type == TokenType.Identifier)
            {
                chain.Add(new ChainSegment(tokens[pos].Lexeme));
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
        // The token before the chain might be Identifier:Identifier pattern
        var hasRangePrefix = false;
        if (pos >= 0 && tokens[pos].Type == TokenType.Colon)
        {
            // Could be part of range ref like A1:B10 — the "B10" is already in chain
            // Check that tokens before : is also an identifier that looks like a cell ref
            if (pos >= 1 && tokens[pos - 1].Type == TokenType.Identifier)
            {
                var startRef = tokens[pos - 1].Lexeme;
                if (chain.Count > 0 && IsCellReference(startRef) && IsCellReference(chain[^1].Name))
                {
                    hasRangePrefix = true;
                    // Remove the end-ref identifier from chain, it's part of the range
                    chain.RemoveAt(chain.Count - 1);
                }
            }
        }

        chain.Reverse();

        if (chain.Count == 0)
        {
            // Just a dot after something we couldn't parse — could be a range ref
            // Check the token directly left of the dot
            if (hasRangePrefix)
            {
                return DslType.Range;
            }

            var leftIdx = dotIndex - 1;
            if (leftIdx >= 0)
            {
                // Check for range ref: Identifier Colon Identifier Dot
                if (tokens[leftIdx].Type == TokenType.Identifier &&
                    leftIdx >= 2 &&
                    tokens[leftIdx - 1].Type == TokenType.Colon &&
                    tokens[leftIdx - 2].Type == TokenType.Identifier &&
                    IsCellReference(tokens[leftIdx - 2].Lexeme) &&
                    IsCellReference(tokens[leftIdx].Lexeme))
                {
                    return DslType.Range;
                }
            }

            return DslType.Unknown;
        }

        // Resolve the root identifier
        var root = chain[0].Name;

        if (hasRangePrefix || IsRangeIdentifier(root, metadata))
        {
            // If hasRangePrefix, the root is the first real member (accessor)
            if (!hasRangePrefix)
            {
                // Root is the range name itself, walk from index 1
                if (chain.Count == 1)
                {
                    return DslType.Range;
                }

                return WalkChain(ApplyMember(DslType.Range, chain[1]), chain, 2);
            }

            return WalkChain(ApplyMember(DslType.Range, chain[0]), chain, 1);
        }

        // Check if root is a lambda parameter
        var lambdaType = ResolveLambdaParamType(tokens, dotIndex, root);
        if (lambdaType != null)
        {
            return WalkChain(lambdaType.Value, chain, 1);
        }

        // Unknown root — try to infer from the chain
        return WalkChain(DslType.Unknown, chain, 1);
    }

    private static DslType WalkChain(DslType type, List<ChainSegment> chain, int startIndex)
    {
        for (var i = startIndex; i < chain.Count; i++)
        {
            type = ApplyMember(type, chain[i]);
        }

        return type;
    }

    private static DslType ApplyMember(DslType current, ChainSegment segment)
    {
        var name = segment.Name;

        return current switch
        {
            DslType.Range when Accessors.Contains(name) => DslType.Pipeline,
            DslType.Range when name.Equals("withHeaders", StringComparison.OrdinalIgnoreCase) => DslType.Range,
            // Ranges default to .values, so pipeline methods work directly on ranges
            DslType.Range when PipelineMethods.Contains(name) => DslType.Pipeline,
            DslType.Range when TerminalMethods.Contains(name) => DslType.Scalar,

            DslType.Pipeline when PipelineMethods.Contains(name) => DslType.Pipeline,
            DslType.Pipeline when TerminalMethods.Contains(name) => DslType.Scalar,

            DslType.Cell when name.Equals("Interior", StringComparison.OrdinalIgnoreCase) => DslType.Interior,
            DslType.Cell when name.Equals("Font", StringComparison.OrdinalIgnoreCase) => DslType.Font,
            DslType.Cell when CellProperties.Contains(name) => DslType.Scalar,

            DslType.Row => DslType.Scalar, // Row.ColumnName is a scalar value

            DslType.Interior => DslType.Scalar, // Interior properties are all scalars
            DslType.Font => DslType.Scalar, // Font properties are all scalars

            // Unknown root + accessor = probably a range we couldn't identify
            DslType.Unknown when Accessors.Contains(name) => DslType.Pipeline,
            DslType.Unknown when PipelineMethods.Contains(name) => DslType.Pipeline,
            DslType.Unknown when TerminalMethods.Contains(name) => DslType.Scalar,

            _ => current // Can't determine further, keep current type
        };
    }

    /// <summary>
    ///     Determines whether a lambda parameter is in a Cell or Row context
    ///     by scanning backward for the enclosing method call.
    ///     Verifies that paramName actually appears as a lambda parameter in the enclosing call.
    /// </summary>
    private static DslType? ResolveLambdaParamType(List<Token> tokens, int dotIndex, string paramName)
    {
        // Scan backward from dotIndex looking for an unmatched `(` that belongs to a
        // method call, then check if paramName is declared as a lambda parameter there.
        var depth = 0; // paren nesting depth

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
                        // Not our lambda — keep scanning outward
                        continue;
                    }

                    // The method name is before this paren
                    if (i >= 2 && tokens[i - 1].Type == TokenType.Identifier &&
                        tokens[i - 2].Type == TokenType.Dot)
                    {
                        var methodName = tokens[i - 1].Lexeme;

                        if (CellContextMethods.Contains(methodName))
                        {
                            return DslType.Cell;
                        }

                        if (RowContextMethods.Contains(methodName))
                        {
                            return DslType.Row;
                        }

                        // For passthrough methods (where, select, etc.), continue scanning
                        // backward to find the accessor that determines cell vs row
                        if (PassthroughMethods.Contains(methodName))
                        {
                            return FindPipelineElementType(tokens, i - 2);
                        }
                    }

                    // Unmatched paren but no recognizable method — keep scanning
                    // (could be grouping parens)
                }
            }
        }

        return null;
    }

    /// <summary>
    ///     Checks whether paramName appears as a lambda parameter between the opening paren
    ///     and an arrow token (=>) within the given range.
    /// </summary>
    private static bool IsLambdaParamInScope(List<Token> tokens, int leftParenIndex, int dotIndex,
        string paramName)
    {
        // Look for paramName before the first => after the opening paren
        for (var j = leftParenIndex + 1; j < dotIndex; j++)
        {
            if (tokens[j].Type == TokenType.Arrow)
            {
                // Found the arrow — check if paramName appeared between paren and arrow
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

        // No arrow found — could be an incomplete lambda; check if paramName is right after paren
        // e.g., `.where(c =>` where we're inside a statement lambda
        return false;
    }

    /// <summary>
    ///     Walks backward from a method's dot position to find whether the pipeline
    ///     operates on cells or rows (by finding the accessor).
    /// </summary>
    private static DslType? FindPipelineElementType(List<Token> tokens, int dotPos)
    {
        // From the dot before the method, continue walking left through the chain
        // looking for the accessor (.cells or .rows)
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
                    var name = tokens[pos].Lexeme;
                    if (CellContextMethods.Contains(name))
                    {
                        return DslType.Cell;
                    }

                    if (RowContextMethods.Contains(name))
                    {
                        return DslType.Row;
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
                var name = tokens[pos].Lexeme;
                if (CellContextMethods.Contains(name))
                {
                    return DslType.Cell;
                }

                if (RowContextMethods.Contains(name))
                {
                    return DslType.Row;
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

        // Check table names
        foreach (var t in metadata.TableNames)
        {
            if (t.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        // Check named ranges (may contain dots — for now check exact match)
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
        // Match patterns like A1, AB12, $A$1 (strip $ signs first)
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

    private record ChainSegment(string Name);
}
