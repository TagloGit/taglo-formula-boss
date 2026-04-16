using System.Diagnostics;

using FormulaBoss.Compilation;
using FormulaBoss.Transpilation;
using FormulaBoss.UI;

namespace FormulaBoss.Interception;

/// <summary>
///     Result of processing a backtick expression through the pipeline.
/// </summary>
/// <param name="Success">Whether processing succeeded.</param>
/// <param name="UdfName">The generated UDF name (if successful).</param>
/// <param name="ErrorMessage">Error message (if failed).</param>
/// <param name="Parameters">Flat list of parameter names for the UDF call, in order.</param>
public record PipelineResult(
    bool Success,
    string? UdfName,
    string? ErrorMessage,
    IReadOnlyList<string>? Parameters = null);

/// <summary>
///     Context for processing a DSL expression, used for LET integration.
/// </summary>
/// <param name="PreferredUdfName">Optional preferred name for the UDF (e.g., from a LET variable).</param>
/// <param name="Metadata">Optional workbook metadata for metadata-aware header detection.</param>
public record ExpressionContext(
    string? PreferredUdfName,
    WorkbookMetadata? Metadata = null);

/// <summary>
///     Orchestrates the complete pipeline: parse → transpile → compile → register.
/// </summary>
public class FormulaPipeline
{
    private readonly DynamicCompiler _compiler;
    private readonly Dictionary<string, IReadOnlyList<string>?> _parametersCache = [];

    // Maps UDF names to the expression they were created from, to detect collisions
    private readonly Dictionary<string, string> _registeredUdfExpressions = [];
    private readonly Dictionary<string, string> _udfCache = [];

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
        // For cache key, include preferred name if provided (same expression with different names = different UDFs)
        var cacheKey = context?.PreferredUdfName != null
            ? $"{expression}|{context.PreferredUdfName}"
            : expression;

        // Check cache first
        if (_udfCache.TryGetValue(cacheKey, out var cachedUdfName))
        {
            _parametersCache.TryGetValue(cacheKey, out var cachedParams);
            return new PipelineResult(true, cachedUdfName, null, cachedParams);
        }

        // Step 1: Detect parameters using Roslyn
        var detector = new InputDetector();
        DetectionResult detection;
        try
        {
            detection = detector.Detect(expression);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Detection error: {ex.Message}");
        }

        // Augment header variables with metadata: a parameter needs [#All] if its name
        // matches a known table OR if the existing AST pattern matching detected it.
        // This must happen before code emission so the generated code extracts headers.
        var metadata = context?.Metadata;
        var headerVariables = detection.Parameters
            .Where(p => detection.HeaderVariables.Contains(p) ||
                        (metadata?.IsTable(p) == true && !detection.RangeRefMap.ContainsKey(p)))
            .ToHashSet();

        // If metadata added new header variables beyond what AST detected, update the
        // detection result so CodeEmitter generates header extraction code for them
        var emitDetection = headerVariables.SetEquals(detection.HeaderVariables)
            ? detection
            : detection with { HeaderVariables = headerVariables };

        // Step 2: Emit code
        var emitter = new CodeEmitter();
        var preferredName = context?.PreferredUdfName;
        if (preferredName != null)
        {
            preferredName = GetUniqueUdfName(preferredName, expression);
        }

        TranspileResult transpileResult;
        try
        {
            transpileResult = emitter.Emit(emitDetection, expression, preferredName);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Emit error: {ex.Message}");
        }

        // Debug: Output the generated source code
        Debug.WriteLine("=== Generated UDF Source Code ===");
        Debug.WriteLine(transpileResult.SourceCode);
        Debug.WriteLine("=== End Generated Code ===");

        // Step 3: Compile and Register
        var compileErrors =
            _compiler.CompileAndRegister(transpileResult.SourceCode, transpileResult.RequiresObjectModel);

        if (compileErrors.Count > 0)
        {
            var errorMsg = string.Join("; ", compileErrors);
            if (ContainsTypeError(errorMsg))
            {
                errorMsg += GetStatementLambdaHint();
            }

            return new PipelineResult(false, null, $"Compile error: {errorMsg}");
        }

        // Step 3b: Emit and compile the debug-instrumented variant alongside the normal one.
        // The caller address expression invokes the delegate bridge so the generated code
        // does not need to reference ExcelDNA's XlCall directly.
        try
        {
            var callerAddrExpr = "FormulaBoss.RuntimeHelpers.GetCallerAddressDelegate?.Invoke() ?? \"\"";
            var debugResult = emitter.EmitDebug(
                emitDetection, expression, preferredName,
                headersByParameter: null, callerAddrExpr);

            Debug.WriteLine("=== Generated DEBUG UDF Source Code ===");
            Debug.WriteLine(debugResult.SourceCode);
            Debug.WriteLine("=== End Generated DEBUG Code ===");

            var debugErrors = _compiler.CompileAndRegister(
                debugResult.SourceCode, debugResult.RequiresObjectModel);

            if (debugErrors.Count > 0)
            {
                Debug.WriteLine($"Debug variant compile failed: {string.Join("; ", debugErrors)}");
            }
            else
            {
                transpileResult = transpileResult with { DebugVariant = debugResult };
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Debug variant emit failed: {ex.Message}");
        }

        // Track which expression this UDF name was created from
        _registeredUdfExpressions[transpileResult.MethodName] = expression;

        // Build flat parameter list, mapping range ref placeholders back to originals
        var parameters = detection.Parameters
            .Select(p =>
            {
                if (detection.RangeRefMap.TryGetValue(p, out var orig))
                {
                    return orig;
                }

                if (headerVariables.Contains(p))
                {
                    return p + "[#All]";
                }

                return p;
            })
            .ToList();

        // Cache the result
        _udfCache[cacheKey] = transpileResult.MethodName;
        _parametersCache[cacheKey] = parameters;

        return new PipelineResult(true, transpileResult.MethodName, null, parameters);
    }

    /// <summary>
    ///     Resolves the UDF name for a preferred name. If the preferred name is already
    ///     registered with a different expression, the old registration is overwritten
    ///     (cache invalidated) so all cells referencing the UDF get the updated behaviour.
    /// </summary>
    private string GetUniqueUdfName(string preferredName, string expression)
    {
        var fullName = FullMethodName(preferredName);

        if (_registeredUdfExpressions.TryGetValue(fullName, out var existingExpression)
            && existingExpression != expression)
        {
            // Different expression wants the same name — overwrite.
            // Invalidate the stale cache entry so the old expression isn't served from cache.
            var oldCacheKey = $"{existingExpression}|{preferredName}";
            _udfCache.Remove(oldCacheKey);
            _parametersCache.Remove(oldCacheKey);

            Debug.WriteLine($"UDF overwrite: {preferredName} re-edited with new expression");
        }

        return preferredName;
    }

    private static string FullMethodName(string preferredName) =>
        CodeEmitter.GenerateMethodName("", preferredName);

    private static bool ContainsTypeError(string errorMsg)
    {
        var typeKeywords = new[]
        {
            "cannot convert", "no implicit conversion", "cannot implicitly convert", "operator",
            "cannot be applied to operands of type", "does not contain a definition for", "cannot be used as",
            "cannot assign"
        };

        return typeKeywords.Any(keyword =>
            errorMsg.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetStatementLambdaHint() =>
        "\n\nHint: Use standard C# casts for type conversion:\n" +
        "  Convert.ToDouble(x) or (double)x\n" +
        "  (string)x or x.ToString()\n" +
        "  (bool)x or Convert.ToBoolean(x)\n" +
        "  (int)x or Convert.ToInt32(x)";
}
