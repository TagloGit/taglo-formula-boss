namespace FormulaBoss.Interception;

/// <summary>
/// Represents a backtick expression found within a formula.
/// </summary>
/// <param name="Expression">The DSL expression inside the backticks.</param>
/// <param name="StartIndex">Start position of the opening backtick in the original text.</param>
/// <param name="EndIndex">End position (exclusive) after the closing backtick.</param>
public record BacktickExpression(string Expression, int StartIndex, int EndIndex);

/// <summary>
/// Extracts backtick-delimited DSL expressions from Excel formulas.
/// </summary>
public static class BacktickExtractor
{
    /// <summary>
    /// Checks if a cell value is a text entry that looks like a formula with backticks.
    /// When user types '=..., Excel stores it as text starting with = (apostrophe is hidden).
    /// </summary>
    /// <param name="cellText">The cell text/value.</param>
    /// <returns>True if this is text starting with = that contains backticks.</returns>
    public static bool IsBacktickFormula(string? cellText)
    {
        if (string.IsNullOrEmpty(cellText))
        {
            return false;
        }

        // Text must start with = (the apostrophe prefix is not visible in cell properties)
        if (!cellText.StartsWith('='))
        {
            return false;
        }

        // Must contain at least one backtick
        return cellText.Contains('`');
    }

    /// <summary>
    /// Extracts all backtick expressions from a formula.
    /// </summary>
    /// <param name="formulaText">The formula text (with or without leading apostrophe).</param>
    /// <returns>List of backtick expressions found.</returns>
    public static List<BacktickExpression> Extract(string formulaText)
    {
        var results = new List<BacktickExpression>();

        var i = 0;
        while (i < formulaText.Length)
        {
            // Find opening backtick
            var start = formulaText.IndexOf('`', i);
            if (start == -1)
            {
                break;
            }

            // Find closing backtick
            var end = formulaText.IndexOf('`', start + 1);
            if (end == -1)
            {
                // Unclosed backtick - could be an error, but we'll ignore it
                break;
            }

            // Extract the expression (without the backticks)
            var expression = formulaText.Substring(start + 1, end - start - 1);
            results.Add(new BacktickExpression(expression, start, end + 1));

            i = end + 1;
        }

        return results;
    }

    /// <summary>
    /// Rewrites a formula by replacing backtick expressions with UDF calls.
    /// </summary>
    /// <param name="originalFormula">The original formula text (starting with =).</param>
    /// <param name="replacements">Dictionary mapping original expressions to UDF call strings.</param>
    /// <returns>The rewritten formula with backtick expressions replaced.</returns>
    public static string RewriteFormula(string originalFormula, Dictionary<string, string> replacements)
    {
        var formula = originalFormula;

        // Replace each backtick expression with its UDF call
        foreach (var (expression, udfCall) in replacements)
        {
            var backtickExpr = $"`{expression}`";
            formula = formula.Replace(backtickExpr, udfCall, StringComparison.Ordinal);
        }

        return formula;
    }
}
