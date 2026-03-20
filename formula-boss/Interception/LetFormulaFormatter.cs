using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
///     Formats LET formulas with configurable indentation, nesting depth, and line length.
///     Preserves non-LET function calls, string literals, array constants, and backtick expressions.
/// </summary>
public static class LetFormulaFormatter
{
    private const char NewLine = '\n'; // Use LF only - Excel COM doesn't like \r\n

    /// <summary>
    ///     Formats a LET formula according to the specified settings.
    ///     Returns the original formula unchanged if it cannot be parsed or formatting is disabled.
    /// </summary>
    /// <param name="formula">The formula to format.</param>
    /// <param name="indentSize">Number of spaces per indent level.</param>
    /// <param name="nestedLetDepth">How many levels of nested LETs to format (0 = off, 1 = top only).</param>
    /// <param name="maxLineLength">Max line length before wrapping (0 = always wrap).</param>
    /// <returns>The formatted formula.</returns>
    public static string Format(string formula, int indentSize = 2, int nestedLetDepth = 1, int maxLineLength = 0)
    {
        if (string.IsNullOrWhiteSpace(formula) || nestedLetDepth <= 0)
        {
            return formula;
        }

        // Check if the formula starts with = followed by optional wrapping function then LET(
        var trimmed = formula.TrimStart();
        if (!trimmed.StartsWith("=", StringComparison.Ordinal))
        {
            return formula;
        }

        // Find the LET( in the formula
        var afterEquals = trimmed.Substring(1).TrimStart();
        if (afterEquals.StartsWith("LET(", StringComparison.OrdinalIgnoreCase))
        {
            // Direct =LET(...) formula
            return FormatDirectLet(trimmed, indentSize, nestedLetDepth, maxLineLength);
        }

        // Check for LET inside another function: =FUNC(LET(...))
        return FormatWrappedLet(trimmed, indentSize, nestedLetDepth, maxLineLength);
    }

    private static string FormatDirectLet(string formula, int indentSize, int nestedLetDepth, int maxLineLength)
    {
        if (!LetFormulaParser.TryParse(formula, out var structure) || structure == null)
        {
            return formula;
        }

        var sb = new StringBuilder();
        sb.Append("=LET(");
        var indent = new string(' ', indentSize);
        FormatLetBody(sb, structure, indent, indentSize, nestedLetDepth, maxLineLength, "=LET(");
        return sb.ToString();
    }

    private static string FormatWrappedLet(string formula, int indentSize, int nestedLetDepth, int maxLineLength)
    {
        // Find pattern like =FUNC(LET(...)...)
        // We need to find "LET(" that's inside another function call
        var letIndex = formula.IndexOf("LET(", 1, StringComparison.OrdinalIgnoreCase);
        if (letIndex == -1)
        {
            return formula;
        }

        // Extract the wrapper prefix (e.g., "=SUMPRODUCT(")
        var prefix = formula.Substring(0, letIndex);

        // Check that the prefix is like "=FUNC(" — ends with a function call opening
        if (!prefix.TrimEnd().EndsWith("("))
        {
            return formula;
        }

        // Parse the inner LET
        var innerLetFormula = "=" + formula.Substring(letIndex);
        if (!LetFormulaParser.TryParse(innerLetFormula, out var structure) || structure == null)
        {
            return formula;
        }

        // Find the closing paren of the LET
        var letOpenParen = letIndex + 3; // index of '(' in "LET("
        var letCloseParen = LetArgumentSplitter.FindMatchingCloseParen(formula, letOpenParen);
        if (letCloseParen == -1)
        {
            return formula;
        }

        // Get any suffix after the LET's closing paren (e.g., "))" for the wrapper)
        var suffix = formula.Substring(letCloseParen + 1);

        var sb = new StringBuilder();
        sb.Append(prefix).Append("LET(");
        var indent = new string(' ', indentSize);
        FormatLetBody(sb, structure, indent, indentSize, nestedLetDepth, maxLineLength, prefix + "LET(");
        sb.Append(suffix);
        return sb.ToString();
    }

    private static void FormatLetBody(
        StringBuilder sb,
        LetStructure structure,
        string indent,
        int indentSize,
        int remainingDepth,
        int maxLineLength,
        string openingLine)
    {
        var bindings = structure.Bindings;
        var result = structure.ResultExpression.Trim();

        // Try inlining with MaxLineLength > 0
        if (maxLineLength > 0)
        {
            FormatWithInlining(sb, bindings, result, indent, indentSize, remainingDepth, maxLineLength, openingLine);
            return;
        }

        // Default: always wrap every binding
        sb.Append(NewLine);
        for (var i = 0; i < bindings.Count; i++)
        {
            var name = bindings[i].VariableName.Trim();
            var value = bindings[i].Value.Trim();
            value = MaybeFormatNestedLet(value, indent, indentSize, remainingDepth);
            sb.Append(indent).Append(name).Append(", ").Append(value).Append(',').Append(NewLine);
        }

        result = MaybeFormatNestedLet(result, indent, indentSize, remainingDepth);
        sb.Append(indent).Append(result).Append(')');
    }

    private static void FormatWithInlining(
        StringBuilder sb,
        List<LetBinding> bindings,
        string result,
        string indent,
        int indentSize,
        int remainingDepth,
        int maxLineLength,
        string openingLine)
    {
        var currentLineLength = openingLine.Length;
        var wrapping = false;
        var firstOnLine = true;

        for (var i = 0; i < bindings.Count; i++)
        {
            var name = bindings[i].VariableName.Trim();
            var value = bindings[i].Value.Trim();
            value = MaybeFormatNestedLet(value, indent, indentSize, remainingDepth);

            // "name, value," is the binding text
            var bindingText = name + ", " + value + ",";

            if (!wrapping)
            {
                // Try to fit on current line
                var spacer = firstOnLine ? "" : " ";
                var wouldBe = currentLineLength + spacer.Length + bindingText.Length;
                if (firstOnLine || wouldBe <= maxLineLength)
                {
                    sb.Append(spacer).Append(bindingText);
                    currentLineLength += spacer.Length + bindingText.Length;
                    firstOnLine = false;
                }
                else
                {
                    // This binding doesn't fit — start wrapping from here
                    wrapping = true;
                    sb.Append(NewLine);
                    sb.Append(indent).Append(bindingText);
                }
            }
            else
            {
                sb.Append(NewLine);
                sb.Append(indent).Append(bindingText);
            }
        }

        // Result expression
        result = MaybeFormatNestedLet(result, indent, indentSize, remainingDepth);
        if (!wrapping)
        {
            var spacer = firstOnLine ? "" : " ";
            sb.Append(spacer).Append(result).Append(')');
        }
        else
        {
            sb.Append(NewLine);
            sb.Append(indent).Append(result).Append(')');
        }
    }

    /// <summary>
    ///     If the value contains a nested LET and we have remaining depth, format it recursively.
    /// </summary>
    private static string MaybeFormatNestedLet(string value, string parentIndent, int indentSize, int remainingDepth)
    {
        if (remainingDepth <= 1)
        {
            return value;
        }

        // Check if the value is a LET(...) expression (possibly with whitespace)
        var trimmed = value.TrimStart();
        if (!trimmed.StartsWith("LET(", StringComparison.OrdinalIgnoreCase))
        {
            return value;
        }

        // Parse the nested LET
        var nestedFormula = "=" + trimmed;
        if (!LetFormulaParser.TryParse(nestedFormula, out var nestedStructure) || nestedStructure == null)
        {
            return value;
        }

        // Format recursively with deeper indent
        var nestedIndent = parentIndent + new string(' ', indentSize);
        var sb = new StringBuilder();
        sb.Append("LET(");
        // For nested LETs, always wrap (maxLineLength=0) to keep things readable
        sb.Append(NewLine);
        for (var i = 0; i < nestedStructure.Bindings.Count; i++)
        {
            var name = nestedStructure.Bindings[i].VariableName.Trim();
            var val = nestedStructure.Bindings[i].Value.Trim();
            val = MaybeFormatNestedLet(val, nestedIndent, indentSize, remainingDepth - 1);
            sb.Append(nestedIndent).Append(name).Append(", ").Append(val).Append(',').Append(NewLine);
        }

        var nestedResult = nestedStructure.ResultExpression.Trim();
        nestedResult = MaybeFormatNestedLet(nestedResult, nestedIndent, indentSize, remainingDepth - 1);
        sb.Append(nestedIndent).Append(nestedResult).Append(')');

        return sb.ToString();
    }
}
