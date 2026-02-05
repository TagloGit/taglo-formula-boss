using System.Text;
using System.Text.RegularExpressions;

namespace FormulaBoss.Interception;

/// <summary>
///     Represents a variable binding in a LET formula.
/// </summary>
/// <param name="VariableName">The name of the LET variable.</param>
/// <param name="Value">The value expression assigned to the variable.</param>
/// <param name="HasBacktick">Whether the value contains a backtick expression.</param>
public record LetBinding(string VariableName, string Value, bool HasBacktick);

/// <summary>
///     Represents a parsed LET formula structure.
/// </summary>
/// <param name="Bindings">The variable bindings (name-value pairs).</param>
/// <param name="ResultExpression">The final result expression of the LET.</param>
/// <param name="OriginalFormula">The original formula text.</param>
public record LetStructure(List<LetBinding> Bindings, string ResultExpression, string OriginalFormula);

/// <summary>
///     Parses Excel LET formulas to extract variable bindings and detect backtick expressions.
/// </summary>
public static class LetFormulaParser
{
    /// <summary>
    ///     Checks if a formula is a LET formula.
    /// </summary>
    /// <param name="formula">The formula text (should start with =).</param>
    /// <returns>True if the formula is a LET statement.</returns>
    public static bool IsLetFormula(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return false;
        }

        // Normalize: trim and check for =LET( pattern (case-insensitive)
        var trimmed = formula.TrimStart();
        return trimmed.StartsWith("=LET(", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Attempts to parse a formula as a LET statement.
    /// </summary>
    /// <param name="formula">The formula text.</param>
    /// <param name="structure">The parsed LET structure if successful.</param>
    /// <returns>True if parsing succeeded.</returns>
    public static bool TryParse(string formula, out LetStructure? structure)
    {
        structure = null;

        if (!IsLetFormula(formula))
        {
            return false;
        }

        try
        {
            structure = Parse(formula);
            return structure != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     Parses a LET formula into its structure.
    /// </summary>
    /// <param name="formula">The formula text.</param>
    /// <returns>The parsed structure, or null if parsing fails.</returns>
    private static LetStructure? Parse(string formula)
    {
        // Find the LET( opening
        var letStart = formula.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letStart == -1)
        {
            return null;
        }

        // Find the matching closing parenthesis
        var bodyStart = letStart + 4; // After "LET("
        var bodyEnd = FindMatchingCloseParen(formula, bodyStart - 1);
        if (bodyEnd == -1)
        {
            return null;
        }

        // Extract the body between LET( and )
        var body = formula.Substring(bodyStart, bodyEnd - bodyStart);

        // Split into arguments respecting nesting
        var arguments = SplitArguments(body);
        if (arguments.Count < 3)
        {
            return null; // LET needs at least: name, value, result
        }

        // LET arguments are pairs: name1, value1, name2, value2, ..., result
        // So we should have an odd number of arguments
        if (arguments.Count % 2 == 0)
        {
            return null; // Invalid: even number means missing result
        }

        var bindings = new List<LetBinding>();
        for (var i = 0; i < arguments.Count - 1; i += 2)
        {
            var variableName = arguments[i].Trim();
            var value = arguments[i + 1].Trim();
            var hasBacktick = value.Contains('`');
            bindings.Add(new LetBinding(variableName, value, hasBacktick));
        }

        var resultExpression = arguments[^1].Trim();

        return new LetStructure(bindings, resultExpression, formula);
    }

    /// <summary>
    ///     Finds the matching closing parenthesis for an opening parenthesis.
    /// </summary>
    /// <param name="text">The text to search.</param>
    /// <param name="openParenIndex">Index of the opening parenthesis.</param>
    /// <returns>Index of the matching closing parenthesis, or -1 if not found.</returns>
    private static int FindMatchingCloseParen(string text, int openParenIndex)
    {
        var depth = 0;
        var inString = false;
        var stringChar = '\0';

        for (var i = openParenIndex; i < text.Length; i++)
        {
            var c = text[i];

            // Handle string literals
            if (!inString && (c == '"' || c == '\''))
            {
                inString = true;
                stringChar = c;
            }
            else if (inString && c == stringChar)
            {
                // Check for escaped quote (doubled)
                if (i + 1 < text.Length && text[i + 1] == stringChar)
                {
                    i++; // Skip the escaped quote
                }
                else
                {
                    inString = false;
                }
            }
            else if (!inString)
            {
                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }
        }

        return -1;
    }

    /// <summary>
    ///     Splits LET arguments by comma, respecting nested parentheses and string literals.
    /// </summary>
    /// <param name="body">The LET body (content between LET( and )).</param>
    /// <returns>List of argument strings.</returns>
    private static List<string> SplitArguments(string body)
    {
        var arguments = new List<string>();
        var current = new StringBuilder();
        var depth = 0;
        var inString = false;
        var stringChar = '\0';

        for (var i = 0; i < body.Length; i++)
        {
            var c = body[i];

            // Handle string literals
            if (!inString && (c == '"' || c == '\''))
            {
                inString = true;
                stringChar = c;
                current.Append(c);
            }
            else if (inString && c == stringChar)
            {
                current.Append(c);
                // Check for escaped quote (doubled)
                if (i + 1 < body.Length && body[i + 1] == stringChar)
                {
                    current.Append(body[i + 1]);
                    i++; // Skip the escaped quote
                }
                else
                {
                    inString = false;
                }
            }
            else if (inString)
            {
                current.Append(c);
            }
            else if (c == '(')
            {
                depth++;
                current.Append(c);
            }
            else if (c == ')')
            {
                depth--;
                current.Append(c);
            }
            else if (c == ',' && depth == 0)
            {
                arguments.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        // Add the last argument
        if (current.Length > 0)
        {
            arguments.Add(current.ToString());
        }

        return arguments;
    }

    /// <summary>
    ///     Extracts the backtick expression from a LET binding value.
    ///     Returns the expression without backticks, or null if no backticks.
    /// </summary>
    /// <param name="value">The binding value.</param>
    /// <returns>The expression inside backticks, or null.</returns>
    public static string? ExtractBacktickExpression(string value)
    {
        var start = value.IndexOf('`');
        if (start == -1)
        {
            return null;
        }

        var end = value.IndexOf('`', start + 1);
        if (end == -1)
        {
            return null;
        }

        return value.Substring(start + 1, end - start - 1);
    }

    /// <summary>
    ///     Checks if a binding value is a table column reference like tblSales[Price].
    /// </summary>
    /// <param name="value">The binding value.</param>
    /// <param name="tableName">The table name if it's a column reference.</param>
    /// <param name="columnName">The column name if it's a column reference.</param>
    /// <returns>True if this is a column reference.</returns>
    public static bool IsColumnBinding(string value, out string? tableName, out string? columnName)
    {
        tableName = null;
        columnName = null;

        // Match pattern: tableName[ColumnName] (with optional spaces)
        // e.g., tblSales[Price], data[Total Amount], Sales[Qty]
        var match = Regex.Match(value.Trim(), @"^(\w+)\s*\[\s*([^\]]+)\s*\]$");
        if (match.Success)
        {
            tableName = match.Groups[1].Value;
            columnName = match.Groups[2].Value.Trim();
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Extracts all column bindings from a LET structure.
    ///     Returns a dictionary mapping LET variable names to column names.
    /// </summary>
    /// <param name="structure">The parsed LET structure.</param>
    /// <returns>Dictionary of variable name → column name.</returns>
    public static Dictionary<string, string> ExtractColumnBindings(LetStructure structure)
    {
        var columnBindings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var binding in structure.Bindings)
        {
            if (IsColumnBinding(binding.Value, out _, out var columnName) && columnName != null)
            {
                columnBindings[binding.VariableName] = columnName;
            }
        }

        return columnBindings;
    }
}
