using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
///     Information about a processed LET binding for rewriting.
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
///     Rewrites LET formulas after backtick expressions have been processed.
///     Inserts _src_* documentation variables and replaces backtick expressions with UDF calls.
/// </summary>
public static class LetFormulaRewriter
{
    /// <summary>
    ///     Rewrites a LET formula, inserting _src_ documentation variables
    ///     and replacing backtick expressions with UDF calls.
    ///     Formats output with one binding per line for readability.
    /// </summary>
    /// <param name="original">The parsed LET structure.</param>
    /// <param name="processedBindings">Backtick bindings that were compiled to UDFs.</param>
    /// <param name="processedResults">Backtick expressions in the result, if any.</param>
    /// <param name="rewrittenResultExpression">Rewritten result expression referencing new bindings.</param>
    /// <param name="indentSize">Number of spaces per indent level.</param>
    /// <param name="nestedLetDepth">How many levels of nested LETs to format (0 = off, 1 = top only).</param>
    /// <param name="maxLineLength">Max line length before wrapping (0 = always wrap).</param>
    public static string Rewrite(
        LetStructure original,
        IReadOnlyDictionary<string, ProcessedBinding> processedBindings,
        IReadOnlyList<ProcessedBinding>? processedResults = null,
        string? rewrittenResultExpression = null,
        int indentSize = 4,
        int nestedLetDepth = 1,
        int maxLineLength = 0)
    {
        // Build a flat (single-line) formula with _src_ bindings inserted,
        // then let LetFormulaFormatter handle all formatting.
        var sb = new StringBuilder();
        sb.Append("=LET(");

        foreach (var binding in original.Bindings)
        {
            var variableName = binding.VariableName.Trim();

            if (processedBindings.TryGetValue(variableName, out var processed))
            {
                // This binding had a backtick expression - insert _src_ and UDF call
                sb.Append("_src_").Append(variableName).Append(", ");
                sb.Append('"').Append(EscapeForExcelString(processed.OriginalExpression)).Append("\", ");

                sb.Append(variableName).Append(", ");
                AppendUdfCall(sb, processed);
                sb.Append(", ");
            }
            else
            {
                // Normal binding - keep as-is
                sb.Append(binding.VariableName.Trim()).Append(", ");
                sb.Append(binding.Value.Trim()).Append(", ");
            }
        }

        // Handle the result expression
        if (processedResults is { Count: > 0 })
        {
            // Result expression had backtick(s) - add _src_ doc and binding for each
            foreach (var processedResult in processedResults)
            {
                sb.Append("_src_").Append(processedResult.VariableName).Append(", ");
                sb.Append('"').Append(EscapeForExcelString(processedResult.OriginalExpression)).Append("\", ");

                sb.Append(processedResult.VariableName).Append(", ");
                AppendUdfCall(sb, processedResult);
                sb.Append(", ");
            }

            // Final expression references the new bindings
            sb.Append(rewrittenResultExpression ?? processedResults[0].VariableName);
        }
        else
        {
            // Result expression is plain - keep as-is
            sb.Append(original.ResultExpression.Trim());
        }

        sb.Append(')');

        // Format the flat formula using LetFormulaFormatter for consistent output.
        // Always format at least depth 1 — the rewriter output should always be readable.
        return LetFormulaFormatter.Format(sb.ToString(), indentSize, Math.Max(1, nestedLetDepth), maxLineLength);
    }

    /// <summary>
    ///     Appends a UDF call with all parameters.
    /// </summary>
    private static void AppendUdfCall(StringBuilder sb, ProcessedBinding processed)
    {
        sb.Append(processed.UdfName).Append('(');

        for (var i = 0; i < processed.Parameters.Count; i++)
        {
            if (i > 0)
            {
                sb.Append(", ");
            }

            sb.Append(processed.Parameters[i]);
        }

        sb.Append(')');
    }

    private static string EscapeForExcelString(string value) => value.Replace("\"", "\"\"");
}
