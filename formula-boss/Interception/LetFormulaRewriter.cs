using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
/// Information about a column binding used in a DSL expression.
/// Used to inject header bindings for dynamic column name resolution.
/// </summary>
/// <param name="LetVariableName">The LET variable name (e.g., "price").</param>
/// <param name="TableName">The Excel Table name (e.g., "tblSales").</param>
/// <param name="ColumnName">The column name in the table (e.g., "Price").</param>
public record ColumnParameter(
    string LetVariableName,
    string TableName,
    string ColumnName);

/// <summary>
/// Information about a processed LET binding for rewriting.
/// </summary>
/// <param name="VariableName">The original LET variable name.</param>
/// <param name="OriginalExpression">The original DSL expression (without backticks).</param>
/// <param name="UdfName">The generated UDF name (e.g., "COLOREDCELLS").</param>
/// <param name="InputParameter">The input parameter for the UDF call (e.g., "data").</param>
/// <param name="ColumnParameters">Column bindings used in the expression that need header injection.</param>
public record ProcessedBinding(
    string VariableName,
    string OriginalExpression,
    string UdfName,
    string InputParameter,
    IReadOnlyList<ColumnParameter>? ColumnParameters = null,
    IReadOnlyList<string>? AdditionalInputs = null,
    IReadOnlyList<string>? FreeVariables = null);

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

        // Collect all column parameters that need header bindings injected
        // Use a set to avoid duplicate header bindings
        var injectedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var binding in original.Bindings)
        {
            var variableName = binding.VariableName.Trim();

            if (processedBindings.TryGetValue(variableName, out var processed))
            {
                // Inject header bindings for column parameters (before _src_ and UDF call)
                InjectHeaderBindings(sb, processed.ColumnParameters, injectedHeaders);

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
            // Inject header bindings for result expression column parameters
            InjectHeaderBindings(sb, processedResult.ColumnParameters, injectedHeaders);

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
    /// Injects header bindings for column parameters that haven't been injected yet.
    /// Each header binding is of the form: _price_hdr, INDEX(tblSales[[#Headers],[Price]],1)
    /// We use INDEX(...,1) to force Excel to return the VALUE of the header cell,
    /// not a reference. Without this, the UDF receives an ExcelReference object
    /// instead of the actual column name string.
    /// </summary>
    private static void InjectHeaderBindings(
        StringBuilder sb,
        IReadOnlyList<ColumnParameter>? columnParameters,
        HashSet<string> injectedHeaders)
    {
        if (columnParameters == null || columnParameters.Count == 0)
        {
            return;
        }

        foreach (var param in columnParameters)
        {
            var headerBindingName = $"_{param.LetVariableName}_hdr";

            // Only inject if we haven't already
            if (injectedHeaders.Add(headerBindingName))
            {
                sb.Append(Indent).Append(headerBindingName).Append(", ");
                // Wrap in INDEX(...,1) to force value evaluation instead of returning a reference
                sb.Append("INDEX(").Append(param.TableName).Append("[[#Headers],[").Append(param.ColumnName).Append("]],1),");
                sb.Append(NewLine);
            }
        }
    }

    /// <summary>
    /// Appends a UDF call with the input parameter and any column header arguments.
    /// </summary>
    private static void AppendUdfCall(StringBuilder sb, ProcessedBinding processed)
    {
        sb.Append(processed.UdfName).Append('(').Append(processed.InputParameter);

        // Additional inputs (for multi-input explicit lambdas like (tbl, maxVal) => ...)
        if (processed.AdditionalInputs != null)
        {
            foreach (var input in processed.AdditionalInputs)
            {
                sb.Append(", ").Append(input);
            }
        }

        // Free variables (references to LET variables used in the expression)
        if (processed.FreeVariables != null)
        {
            foreach (var freeVar in processed.FreeVariables)
            {
                sb.Append(", ").Append(freeVar);
            }
        }

        if (processed.ColumnParameters != null && processed.ColumnParameters.Count > 0)
        {
            foreach (var param in processed.ColumnParameters)
            {
                sb.Append(", _").Append(param.LetVariableName).Append("_hdr");
            }
        }

        sb.Append(')');
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
