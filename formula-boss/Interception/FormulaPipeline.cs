using System.Diagnostics;

using FormulaBoss.Compilation;
using FormulaBoss.Transpilation;

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
public record ExpressionContext(
    string? PreferredUdfName);

/// <summary>
///     Orchestrates the complete pipeline: parse → transpile → compile → register.
/// </summary>
public class FormulaPipeline
{
    private readonly DynamicCompiler _compiler;
    private readonly Dictionary<string, IReadOnlyList<string>?> _parametersCache = new();

    // Maps UDF names to the expression they were created from, to detect collisions
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

        // Step 2: Emit code
        var emitter = new CodeEmitter();
        TranspileResult transpileResult;
        try
        {
            var preferredName = context?.PreferredUdfName;

            // If we have a preferred name, check if it's already registered with a different expression
            if (preferredName != null)
            {
                preferredName = GetUniqueUdfName(preferredName, expression);
            }

            transpileResult = emitter.Emit(detection, expression, preferredName);
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
        var compileErrors = _compiler.CompileAndRegister(transpileResult.SourceCode, transpileResult.RequiresObjectModel);

        if (compileErrors.Count > 0)
        {
            var errorMsg = string.Join("; ", compileErrors);
            if (ContainsTypeError(errorMsg))
            {
                errorMsg += GetStatementLambdaHint();
            }
            return new PipelineResult(false, null, $"Compile error: {errorMsg}");
        }

        // Track which expression this UDF name was created from
        _registeredUdfExpressions[transpileResult.MethodName] = expression;

        // Build flat parameter list, mapping range ref placeholders back to originals
        var parameters = detection.Parameters
            .Select(p => detection.RangeRefMap.TryGetValue(p, out var orig) ? orig : p)
            .ToList();

        // Cache the result
        _udfCache[cacheKey] = transpileResult.MethodName;
        _parametersCache[cacheKey] = parameters;

        return new PipelineResult(true, transpileResult.MethodName, null, parameters);
    }

    /// <summary>
    ///     Gets a unique UDF name, appending a suffix if the preferred name is already taken by a different expression.
    /// </summary>
    private string GetUniqueUdfName(string preferredName, string expression)
    {
        var candidateName = preferredName;
        var suffix = 2;

        while (_registeredUdfExpressions.TryGetValue(SanitizeName(candidateName), out var existingExpression))
        {
            // If same expression, we can reuse the name (will hit cache anyway)
            if (existingExpression == expression)
            {
                break;
            }

            // Different expression wants the same name - generate a unique one
            candidateName = $"{preferredName}_{suffix}";
            suffix++;

            Debug.WriteLine($"UDF name collision: {preferredName} already registered, trying {candidateName}");
        }

        return candidateName;
    }

    private static string SanitizeName(string name) => CodeEmitter.SanitizeName(name);

    private static bool ContainsTypeError(string errorMsg)
    {
        var typeKeywords = new[]
        {
            "cannot convert",
            "no implicit conversion",
            "cannot implicitly convert",
            "operator",
            "cannot be applied to operands of type",
            "does not contain a definition for",
            "cannot be used as",
            "cannot assign"
        };

        return typeKeywords.Any(keyword =>
            errorMsg.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetStatementLambdaHint() =>
        "\n\nHint: In statement lambdas, use helper methods for type conversion:\n" +
        "  Num(x) - convert to double\n" +
        "  Str(x) - convert to string\n" +
        "  Bool(x) - convert to bool\n" +
        "  Int(x) - convert to int\n" +
        "  IsEmpty(x) - check if null/empty";
}
