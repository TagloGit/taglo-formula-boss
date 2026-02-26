using System.Diagnostics;
using System.Text;

using FormulaBoss.Compilation;
using FormulaBoss.Transpilation;

namespace FormulaBoss.Interception;

/// <summary>
///     Result of processing a backtick expression through the pipeline.
/// </summary>
/// <param name="Success">Whether processing succeeded.</param>
/// <param name="UdfName">The generated UDF name (if successful).</param>
/// <param name="ErrorMessage">Error message (if failed).</param>
/// <param name="InputParameter">The input parameter name extracted from the expression.</param>
/// <param name="ColumnParameters">Column bindings used in the expression that need header injection.</param>
public record PipelineResult(
    bool Success,
    string? UdfName,
    string? ErrorMessage,
    string? InputParameter,
    IReadOnlyList<ColumnParameter>? ColumnParameters = null);

/// <summary>
///     Context for processing a DSL expression, used for LET integration.
/// </summary>
/// <param name="PreferredUdfName">Optional preferred name for the UDF (e.g., from a LET variable).</param>
/// <param name="ColumnBindings">
///     Optional column bindings from LET variables.
///     Maps LET variable names to column binding info (table and column names).
///     Used to resolve r.price to r[__GetCol__("Price")] and generate dynamic column parameters.
/// </param>
public record ExpressionContext(
    string? PreferredUdfName,
    Dictionary<string, ColumnBindingInfo>? ColumnBindings = null);

/// <summary>
///     Orchestrates the complete pipeline: detect inputs → emit code → compile → register.
/// </summary>
public class FormulaPipeline
{
    private readonly Dictionary<string, IReadOnlyList<ColumnParameter>?> _columnParamsCache = new();
    private readonly DynamicCompiler _compiler;
    private readonly Dictionary<string, string> _registeredUdfExpressions = new();
    private readonly Dictionary<string, string> _udfCache = new();

    public FormulaPipeline(DynamicCompiler compiler)
    {
        _compiler = compiler;
    }

    /// <summary>
    ///     Processes a DSL expression and returns the UDF name to use.
    /// </summary>
    public PipelineResult Process(string expression) => Process(expression, null);

    /// <summary>
    ///     Processes a DSL expression with optional context for LET integration.
    /// </summary>
    public PipelineResult Process(string expression, ExpressionContext? context)
    {
        var cacheKey = context?.PreferredUdfName != null
            ? $"{expression}|{context.PreferredUdfName}"
            : expression;

        if (_udfCache.TryGetValue(cacheKey, out var cachedUdfName))
        {
            var inputParam = ExtractInputParameter(expression);
            _columnParamsCache.TryGetValue(cacheKey, out var cachedColumnParams);
            return new PipelineResult(true, cachedUdfName, null, inputParam, cachedColumnParams);
        }

        // Step 1: Detect inputs using Roslyn
        InputDetectionResult detection;
        try
        {
            var knownLetVars = context?.ColumnBindings?.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
            detection = InputDetector.Detect(expression, knownLetVars);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Input detection error: {ex.Message}", null);
        }

        // Step 2: Emit C# code
        TranspileResult transpileResult;
        try
        {
            var preferredName = context?.PreferredUdfName;
            if (preferredName != null)
            {
                preferredName = GetUniqueUdfName(preferredName, expression);
            }

            var methodName = preferredName ?? GenerateMethodName(expression);
            transpileResult = CodeEmitter.Emit(detection, methodName, expression);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Code emission error: {ex.Message}", null);
        }

        Debug.WriteLine("=== Generated UDF Source Code ===");
        Debug.WriteLine(transpileResult.SourceCode);
        Debug.WriteLine("=== End Generated Code ===");

        // Step 3: Compile and register
        var compileErrors = _compiler.CompileAndRegister(transpileResult.SourceCode, transpileResult.RequiresObjectModel);
        if (compileErrors.Count > 0)
        {
            var errorMsg = string.Join("; ", compileErrors);
            return new PipelineResult(false, null, $"Compile error: {errorMsg}", null);
        }

        _registeredUdfExpressions[transpileResult.MethodName] = expression;

        // Build column parameters from used column bindings
        IReadOnlyList<ColumnParameter>? columnParameters = null;
        if (transpileResult.UsedColumnBindings != null
            && transpileResult.UsedColumnBindings.Count > 0
            && context?.ColumnBindings != null)
        {
            var colParams = new List<ColumnParameter>();
            foreach (var usedBinding in transpileResult.UsedColumnBindings)
            {
                if (context.ColumnBindings.TryGetValue(usedBinding, out var bindingInfo))
                {
                    colParams.Add(new ColumnParameter(usedBinding, bindingInfo.TableName, bindingInfo.ColumnName));
                }
            }

            if (colParams.Count > 0)
            {
                columnParameters = colParams;
            }
        }

        _udfCache[cacheKey] = transpileResult.MethodName;
        _columnParamsCache[cacheKey] = columnParameters;

        var inputParameter = detection.Inputs.Count > 0 ? detection.Inputs[0] : null;

        return new PipelineResult(true, transpileResult.MethodName, null, inputParameter, columnParameters);
    }

    /// <summary>
    ///     Generates a method name from the expression when no preferred name is given.
    ///     Uses the primary input name or a hash-based fallback.
    /// </summary>
    private static string GenerateMethodName(string expression)
    {
        var dotIdx = expression.IndexOf('.');
        if (dotIdx > 0)
        {
            return expression[..dotIdx].Trim();
        }

        // Fallback: hash-based name
        var hash = Math.Abs(expression.GetHashCode()).ToString("X8");
        return $"_UDF_{hash}";
    }

    private string GetUniqueUdfName(string preferredName, string expression)
    {
        var candidateName = preferredName;
        var suffix = 2;

        while (_registeredUdfExpressions.TryGetValue(SanitizeName(candidateName), out var existingExpression))
        {
            if (existingExpression == expression)
            {
                break;
            }

            candidateName = $"{preferredName}_{suffix}";
            suffix++;

            Debug.WriteLine($"UDF name collision: {preferredName} already registered, trying {candidateName}");
        }

        return candidateName;
    }

    private static string SanitizeName(string name)
    {
        var upper = name.ToUpperInvariant();
        var result = new StringBuilder();

        foreach (var c in upper)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                result.Append(c);
            }
        }

        var str = result.ToString();
        if (str.Length == 0)
        {
            return "_UDF";
        }

        if (char.IsDigit(str[0]))
        {
            str = "_" + str;
        }

        if (CodeEmitter.IsReservedExcelName(str))
        {
            str = "_" + str;
        }

        return str;
    }

    /// <summary>
    ///     Extracts the input parameter from a DSL expression.
    /// </summary>
    private static string ExtractInputParameter(string expression)
    {
        var dotIndex = expression.IndexOf('.');
        return dotIndex > 0 ? expression[..dotIndex] : expression;
    }
}
