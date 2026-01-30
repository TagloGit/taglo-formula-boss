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
    /// <summary>
    /// Rewrites a LET formula, inserting _src_ documentation variables
    /// and replacing backtick expressions with UDF calls.
    /// </summary>
    /// <param name="original">The parsed LET structure.</param>
    /// <param name="processedBindings">Dictionary of variable name to processed binding info.</param>
    /// <returns>The rewritten formula.</returns>
    public static string Rewrite(
        LetStructure original,
        IReadOnlyDictionary<string, ProcessedBinding> processedBindings)
    {
        var sb = new StringBuilder();
        sb.Append("=LET(");

        var isFirst = true;

        foreach (var binding in original.Bindings)
        {
            if (!isFirst)
            {
                sb.Append(',');
            }

            isFirst = false;

            if (processedBindings.TryGetValue(binding.VariableName.Trim(), out var processed))
            {
                // This binding had a backtick expression - insert _src_ and UDF call
                // Format: _src_varName, "expression", varName, UDFNAME(input)
                sb.Append($" _src_{binding.VariableName.Trim()}, \"{EscapeForExcelString(processed.OriginalExpression)}\"");
                sb.Append($", {binding.VariableName.Trim()}, {processed.UdfName}({processed.InputParameter})");
            }
            else
            {
                // Normal binding - keep as-is
                sb.Append($" {binding.VariableName}, {binding.Value}");
            }
        }

        // Add the result expression
        sb.Append($", {original.ResultExpression})");

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
