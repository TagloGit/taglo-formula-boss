using System.Text;

namespace FormulaBoss.Interception;

/// <summary>
/// Reconstructs editable backtick formulas from processed Formula Boss LET formulas.
/// </summary>
public static class LetFormulaReconstructor
{
    private const string SourcePrefix = "_src_";
    private const string HeaderSuffix = "_hdr"; // Header bindings for dynamic column names

    /// <summary>
    /// Checks if a formula is a processed Formula Boss LET formula (contains _src_ variables).
    /// </summary>
    /// <param name="formula">The formula to check.</param>
    /// <returns>True if the formula contains _src_ documentation variables.</returns>
    public static bool IsProcessedFormulaBossLet(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return false;
        }

        // Must be a LET formula with _src_ bindings
        if (!LetFormulaParser.TryParse(formula, out var structure) || structure == null)
        {
            return false;
        }

        return structure.Bindings.Any(b => b.VariableName.Trim().StartsWith(SourcePrefix, StringComparison.Ordinal));
    }

    /// <summary>
    /// Attempts to reconstruct the original editable formula from a processed LET formula.
    /// </summary>
    /// <param name="formula">The processed formula (with _src_ variables and UDF calls).</param>
    /// <param name="editableFormula">The reconstructed editable formula with backticks and quote prefix.</param>
    /// <returns>True if reconstruction succeeded; false if not a processed Formula Boss formula.</returns>
    public static bool TryReconstruct(string formula, out string? editableFormula)
    {
        editableFormula = null;

        if (!LetFormulaParser.TryParse(formula, out var structure) || structure == null)
        {
            return false;
        }

        // Build a map of variable name -> DSL expression from _src_ bindings
        var sourceExpressions = new Dictionary<string, string>();

        foreach (var binding in structure.Bindings)
        {
            var varName = binding.VariableName.Trim();
            if (varName.StartsWith(SourcePrefix, StringComparison.Ordinal))
            {
                var targetVarName = varName[SourcePrefix.Length..];
                var dslExpression = UnescapeExcelString(binding.Value.Trim());
                sourceExpressions[targetVarName] = dslExpression;
            }
        }

        // If no _src_ bindings found, this isn't a Formula Boss formula
        if (sourceExpressions.Count == 0)
        {
            return false;
        }

        // Reconstruct as a flat formula — formatting is handled by the caller
        // (FormatLetIfEnabled in DetectEditorContent) so it respects AutoFormatLet.
        var sb = new StringBuilder();
        sb.Append("'=LET(");

        foreach (var binding in structure.Bindings)
        {
            var varName = binding.VariableName.Trim();

            // Skip _src_ bindings - they become backtick expressions
            if (varName.StartsWith(SourcePrefix, StringComparison.Ordinal))
            {
                continue;
            }

            // Skip _*_hdr header bindings - these are injected machinery for dynamic column names
            if (varName.StartsWith("_", StringComparison.Ordinal) &&
                varName.EndsWith(HeaderSuffix, StringComparison.Ordinal))
            {
                continue;
            }

            sb.Append(varName);
            sb.Append(", ");

            // Check if this binding has a corresponding _src_ expression
            if (sourceExpressions.TryGetValue(varName, out var dslExpression))
            {
                // Replace UDF call with backtick expression
                sb.Append('`');
                sb.Append(dslExpression);
                sb.Append('`');
            }
            else
            {
                // Keep the original value
                sb.Append(binding.Value.Trim());
            }

            sb.Append(", ");
        }

        // Handle result expression - emit as-is (variable references like _result
        // are just normal LET bindings, same as any user-chosen name)
        sb.Append(structure.ResultExpression.Trim());

        sb.Append(')');

        editableFormula = sb.ToString();
        return true;
    }

    /// <summary>
    /// Unescapes an Excel string literal value (removes surrounding quotes and unescapes doubled quotes).
    /// </summary>
    /// <param name="value">The string value from the LET binding (may include quotes).</param>
    /// <returns>The unescaped string content.</returns>
    private static string UnescapeExcelString(string value)
    {
        // Remove surrounding quotes if present
        if (value.StartsWith('"') && value.EndsWith('"') && value.Length >= 2)
        {
            value = value[1..^1];
        }

        // Unescape doubled quotes: "" -> "
        return value.Replace("\"\"", "\"");
    }
}
