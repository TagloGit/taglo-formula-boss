using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
/// Information about a processed LET binding for rewriting.
/// </summary>
/// <param name="VariableName">The original LET variable name.</param>
/// <param name="OriginalExpression">The original DSL expression (without backticks).</param>
/// <param name="UdfName">The generated UDF name (e.g., "COLOREDCELLS").</param>
/// <param name="Parameters">Flat list of parameter names for the UDF call.</param>
public record ProcessedBinding(
    string VariableName,
    string OriginalExpression,
    string UdfName,
    IReadOnlyList<string> Parameters);

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
                AppendUdfCall(sb, processed);
                sb.Append(',').Append(NewLine);
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
            AppendUdfCall(sb, processedResult);
            sb.Append(',').Append(NewLine);

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
    /// Appends a UDF call with all parameters.
    /// </summary>
    private static void AppendUdfCall(StringBuilder sb, ProcessedBinding processed)
    {
        sb.Append(processed.UdfName).Append('(');

        for (var i = 0; i < processed.Parameters.Count; i++)
        {
            if (i > 0) sb.Append(", ");
            sb.Append(processed.Parameters[i]);
        }

        sb.Append(')');
    }

    private static string EscapeForExcelString(string value)
    {
        return value.Replace("\"", "\"\"");
    }
}
