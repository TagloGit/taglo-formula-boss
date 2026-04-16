using System.Text;
using System.Text.RegularExpressions;

using FormulaBoss.Transpilation;
using FormulaBoss.UI;

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

        // Build a flat formula, then optionally format using user's settings.
        var sb = new StringBuilder();
        sb.Append("=LET(");

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

        // Format if AutoFormatLet is enabled; otherwise return flat
        var settings = EditorSettings.Load();
        var flat = sb.ToString();
        var result = settings.AutoFormatLet
            ? LetFormulaFormatter.Format(flat, settings.IndentSize, Math.Max(1, settings.NestedLetDepth))
            : flat;

        // Prepend quote prefix for text storage
        editableFormula = "'" + result;
        return true;
    }

    /// <summary>
    /// Returns the names of any UDFs currently using _DEBUG call sites in the formula.
    /// Names are returned without the <c>__FB_</c> prefix and <c>_DEBUG</c> suffix
    /// (e.g. "FILTERED" for a call site <c>__FB_FILTERED_DEBUG(...)</c>).
    /// </summary>
    public static List<string> GetDebugCallSites(string? formula)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(formula))
        {
            return result;
        }

        var pattern = CodeEmitter.UdfPrefix + @"(\w+)" + CodeEmitter.DebugSuffix + @"\(";
        foreach (var segment in EnumerateNonStringSegments(formula))
        {
            if (segment.IsStringLiteral)
            {
                continue;
            }

            foreach (Match match in Regex.Matches(segment.Text, pattern))
            {
                result.Add(match.Groups[1].Value);
            }
        }

        return result;
    }

    /// <summary>
    /// Rewrites call sites from normal to debug: <c>__FB_&lt;NAME&gt;(</c> → <c>__FB_&lt;NAME&gt;_DEBUG(</c>.
    /// Only rewrites the specified names. String literals (e.g. <c>_src_</c> bindings) are not touched.
    /// </summary>
    public static string RewriteCallSitesToDebug(string formula, IEnumerable<string> names)
    {
        var result = formula;
        foreach (var name in names)
        {
            var normalCallSite = CodeEmitter.UdfPrefix + name + "(";
            var debugCallSite = CodeEmitter.UdfPrefix + name + CodeEmitter.DebugSuffix + "(";
            result = ReplaceOutsideStringLiterals(result, normalCallSite, debugCallSite);
        }

        return result;
    }

    /// <summary>
    /// Rewrites call sites from debug to normal: <c>__FB_&lt;NAME&gt;_DEBUG(</c> → <c>__FB_&lt;NAME&gt;(</c>.
    /// Only rewrites the specified names. String literals (e.g. <c>_src_</c> bindings) are not touched.
    /// </summary>
    public static string RewriteCallSitesToNormal(string formula, IEnumerable<string> names)
    {
        var result = formula;
        foreach (var name in names)
        {
            var debugCallSite = CodeEmitter.UdfPrefix + name + CodeEmitter.DebugSuffix + "(";
            var normalCallSite = CodeEmitter.UdfPrefix + name + "(";
            result = ReplaceOutsideStringLiterals(result, debugCallSite, normalCallSite);
        }

        return result;
    }

    /// <summary>
    /// Replaces all occurrences of <paramref name="oldValue"/> with <paramref name="newValue"/>
    /// but only outside of Excel string literals (double-quoted regions).
    /// </summary>
    private static string ReplaceOutsideStringLiterals(string formula, string oldValue, string newValue)
    {
        var sb = new StringBuilder(formula.Length);

        foreach (var segment in EnumerateNonStringSegments(formula))
        {
            if (segment.IsStringLiteral)
            {
                sb.Append(segment.Text);
            }
            else
            {
                sb.Append(segment.Text.Replace(oldValue, newValue, StringComparison.OrdinalIgnoreCase));
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Splits a formula into segments, distinguishing between string literals (double-quoted)
    /// and non-string regions. Handles escaped quotes (<c>""</c>) inside strings.
    /// </summary>
    private static IEnumerable<FormulaSegment> EnumerateNonStringSegments(string formula)
    {
        var i = 0;
        var segmentStart = 0;

        while (i < formula.Length)
        {
            if (formula[i] == '"')
            {
                // Yield any non-string segment before this quote
                if (i > segmentStart)
                {
                    yield return new FormulaSegment(formula[segmentStart..i], false);
                }

                // Find the end of the string literal (handling "" escapes)
                var stringStart = i;
                i++; // skip opening quote
                while (i < formula.Length)
                {
                    if (formula[i] == '"')
                    {
                        i++;
                        // "" is an escaped quote inside the string, not the end
                        if (i < formula.Length && formula[i] == '"')
                        {
                            i++;
                            continue;
                        }

                        // End of string literal
                        break;
                    }

                    i++;
                }

                yield return new FormulaSegment(formula[stringStart..i], true);
                segmentStart = i;
            }
            else
            {
                i++;
            }
        }

        // Yield any remaining non-string segment
        if (segmentStart < formula.Length)
        {
            yield return new FormulaSegment(formula[segmentStart..], false);
        }
    }

    private readonly record struct FormulaSegment(string Text, bool IsStringLiteral);

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
