using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
/// Information about a processed LET binding for rewriting.
/// </summary>
/// <param name="VariableName">The original LET variable name.</param>
/// <param name="OriginalExpression">The original DSL expression (without backticks).</param>
/// <param name="UdfName">The generated UDF name (e.g., "COLOREDCELLS").</param>
/// <param name="InputParameter">The input parameter for the UDF call (e.g., "data").</param>
public record ProcessedBinding(
    string VariableName,
    string OriginalExpression,
    string UdfName,
    string InputParameter);

/// <summary>
/// Rewrites LET formulas after backtick expressions have been processed.
/// Inserts _src_* documentation variables and replaces backtick expressions with UDF calls.
/// </summary>
public static class LetFormulaRewriter
{
    private const string Indent = "    ";
    private const char NewLine = '\n'; // Use LF only - Excel COM doesn't like \r\n

    /// <summary>
    /// Rewrites a LET formula, inserting _src_ documentation variables
    /// and replacing backtick expressions with UDF calls.
    /// Formats output with one binding per line for readability.
    /// </summary>
    /// <param name="original">The parsed LET structure.</param>
    /// <param name="processedBindings">Dictionary of variable name to processed binding info.</param>
    /// <param name="processedResult">Optional processed binding for the result expression if it contained a backtick.</param>
    /// <returns>The rewritten formula.</returns>
    public static string Rewrite(
        LetStructure original,
        IReadOnlyDictionary<string, ProcessedBinding> processedBindings,
        ProcessedBinding? processedResult = null)
    {
        var sb = new StringBuilder();
        sb.Append("=LET(").Append(NewLine);

        foreach (var binding in original.Bindings)
        {
            var variableName = binding.VariableName.Trim();

            if (processedBindings.TryGetValue(variableName, out var processed))
            {
                // This binding had a backtick expression - insert _src_ and UDF call
                sb.Append(Indent).Append("_src_").Append(variableName).Append(", ");
                sb.Append('"').Append(EscapeForExcelString(processed.OriginalExpression)).Append("\",").Append(NewLine);

                sb.Append(Indent).Append(variableName).Append(", ");
                sb.Append(processed.UdfName).Append('(').Append(processed.InputParameter).Append("),").Append(NewLine);
            }
            else
            {
                // Normal binding - keep as-is
                sb.Append(Indent).Append(binding.VariableName).Append(", ");
                sb.Append(binding.Value).Append(',').Append(NewLine);
            }
        }

        // Handle the result expression
        if (processedResult != null)
        {
            // Result expression had a backtick - add _src_ doc and binding, then reference it
            sb.Append(Indent).Append("_src_").Append(processedResult.VariableName).Append(", ");
            sb.Append('"').Append(EscapeForExcelString(processedResult.OriginalExpression)).Append("\",").Append(NewLine);

            sb.Append(Indent).Append(processedResult.VariableName).Append(", ");
            sb.Append(processedResult.UdfName).Append('(').Append(processedResult.InputParameter).Append("),").Append(NewLine);

            // Final expression references the new binding
            sb.Append(Indent).Append(processedResult.VariableName).Append(')');
        }
        else
        {
            // Result expression is plain - keep as-is (no trailing comma)
            sb.Append(Indent).Append(original.ResultExpression.Trim()).Append(')');
        }

        return sb.ToString();
    }

    /// <summary>
    /// Escapes a string for use inside an Excel string literal.
    /// Doubles any quotes to escape them.
    /// </summary>
    private static string EscapeForExcelString(string value)
    {
        // Excel uses doubled quotes to escape: "hello ""world""" = hello "world"
        return value.Replace("\"", "\"\"");
    }
}
