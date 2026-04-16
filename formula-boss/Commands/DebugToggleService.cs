using System.Diagnostics;

using ExcelDna.Integration;

using FormulaBoss.Interception;
using FormulaBoss.Transpilation;

namespace FormulaBoss.Commands;

/// <summary>
///     Orchestrates toggling debug mode for the active cell's formula.
///     Reads the cell formula, flips call sites between normal and _DEBUG variants,
///     ensures the debug UDF is compiled, and writes the updated formula back.
/// </summary>
public class DebugToggleService
{
    private readonly FormulaPipeline _pipeline;

    public DebugToggleService(FormulaPipeline pipeline)
    {
        _pipeline = pipeline;
    }

    /// <summary>
    ///     Toggles debug mode for the given cell. Returns the new debug state.
    /// </summary>
    /// <param name="cell">The Excel cell (COM object).</param>
    /// <returns>True if debug mode is now ON; false if OFF.</returns>
    public bool Toggle(dynamic cell)
    {
        var formula = cell.Formula2 as string ?? cell.Formula as string ?? "";
        if (string.IsNullOrEmpty(formula))
        {
            return false;
        }

        var debugNames = LetFormulaReconstructor.GetDebugCallSites(formula);

        if (debugNames.Count > 0)
        {
            // Currently in debug mode — toggle OFF
            var normalFormula = LetFormulaReconstructor.RewriteCallSitesToNormal(formula, debugNames);
            cell.Formula2 = normalFormula;
            Debug.WriteLine($"Debug toggle OFF: rewrote {debugNames.Count} call sites");
            return false;
        }

        // Currently in normal mode — toggle ON
        var normalNames = GetNormalCallSiteNames(formula);
        if (normalNames.Count == 0)
        {
            Debug.WriteLine("Debug toggle: no FB call sites found");
            return false;
        }

        // Ensure debug variants are compiled for each name
        EnsureDebugVariantsCompiled(formula, normalNames);

        var debugFormula = LetFormulaReconstructor.RewriteCallSitesToDebug(formula, normalNames);
        cell.Formula2 = debugFormula;
        Debug.WriteLine($"Debug toggle ON: rewrote {normalNames.Count} call sites");
        return true;
    }

    /// <summary>
    ///     Checks whether the given formula currently has any _DEBUG call sites.
    /// </summary>
    public static bool IsDebugMode(string? formula)
    {
        return LetFormulaReconstructor.GetDebugCallSites(formula).Count > 0;
    }

    /// <summary>
    ///     Gets the names of normal (non-debug) FB call sites in the formula.
    /// </summary>
    private static List<string> GetNormalCallSiteNames(string formula)
    {
        var names = new List<string>();
        var prefix = CodeEmitter.UdfPrefix;
        var debugSuffix = CodeEmitter.DebugSuffix;

        var searchFrom = 0;
        while (true)
        {
            var idx = formula.IndexOf(prefix, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                break;
            }

            var nameStart = idx + prefix.Length;
            var parenIdx = formula.IndexOf('(', nameStart);
            if (parenIdx < 0)
            {
                break;
            }

            var name = formula[nameStart..parenIdx];

            // Skip if this is already a _DEBUG call site
            if (!name.EndsWith(debugSuffix, StringComparison.OrdinalIgnoreCase))
            {
                names.Add(name);
            }

            searchFrom = parenIdx + 1;
        }

        return names;
    }

    /// <summary>
    ///     Ensures the debug variant is compiled for each named UDF by re-processing
    ///     the DSL source from the formula's _src_ bindings through the pipeline.
    /// </summary>
    private void EnsureDebugVariantsCompiled(string formula, List<string> names)
    {
        if (!LetFormulaParser.TryParse(formula, out var structure) || structure == null)
        {
            return;
        }

        // Build a map of variable name -> DSL source from _src_ bindings
        var sourceMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var binding in structure.Bindings)
        {
            var varName = binding.VariableName.Trim();
            if (varName.StartsWith("_src_", StringComparison.Ordinal))
            {
                var targetName = varName["_src_".Length..];
                var value = binding.Value.Trim();
                // Unescape the Excel string literal
                if (value.StartsWith('"') && value.EndsWith('"') && value.Length >= 2)
                {
                    value = value[1..^1].Replace("\"\"", "\"");
                }

                sourceMap[targetName] = value;
            }
        }

        // Re-process each name through the pipeline to ensure debug variants exist.
        // The pipeline compiles both normal and debug variants, and caching prevents
        // redundant compilation if they're already registered.
        foreach (var name in names)
        {
            if (sourceMap.TryGetValue(name, out var source))
            {
                try
                {
                    var context = new ExpressionContext(name);
                    _pipeline.Process(source, context);
                    Debug.WriteLine($"Debug variant ensured for: {name}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to compile debug variant for {name}: {ex.Message}");
                }
            }
        }
    }
}
