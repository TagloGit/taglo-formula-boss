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
///     Orchestrates the complete pipeline: parse → transpile → compile → register.
/// </summary>
public class FormulaPipeline
{
    // Cache for column parameters and input parameters alongside UDF names
    private readonly Dictionary<string, IReadOnlyList<ColumnParameter>?> _columnParamsCache = new();
    private readonly DynamicCompiler _compiler;
    private readonly Dictionary<string, string?> _inputParamCache = new();

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
    /// <param name="expression">The DSL expression (without backticks).</param>
    /// <returns>The pipeline result.</returns>
    public PipelineResult Process(string expression) => Process(expression, null);

    /// <summary>
    ///     Processes a DSL expression with optional context for LET integration.
    /// </summary>
    /// <param name="expression">The DSL expression (without backticks).</param>
    /// <param name="context">Optional context containing preferred UDF name and known variables.</param>
    /// <returns>The pipeline result.</returns>
    public PipelineResult Process(string expression, ExpressionContext? context)
    {
        // For cache key, include preferred name if provided (same expression with different names = different UDFs)
        var cacheKey = context?.PreferredUdfName != null
            ? $"{expression}|{context.PreferredUdfName}"
            : expression;

        // Check cache first
        if (_udfCache.TryGetValue(cacheKey, out var cachedUdfName))
        {
            _inputParamCache.TryGetValue(cacheKey, out var cachedInputParam);
            _columnParamsCache.TryGetValue(cacheKey, out var cachedColumnParams);
            return new PipelineResult(true, cachedUdfName, null, cachedInputParam, cachedColumnParams);
        }

        // Step 1: Detect inputs using Roslyn
        var detector = new InputDetector();
        DetectionResult detection;
        try
        {
            detection = detector.Detect(expression);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Detection error: {ex.Message}", null);
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
            return new PipelineResult(false, null, $"Emit error: {ex.Message}", null);
        }

        // Debug: Output the generated source code
        Debug.WriteLine("=== Generated UDF Source Code ===");
        Debug.WriteLine(transpileResult.SourceCode);
        Debug.WriteLine("=== End Generated Code ===");

        // Step 4: Compile and Register
        var compileErrors = _compiler.CompileAndRegister(transpileResult.SourceCode, transpileResult.RequiresObjectModel);

        if (compileErrors.Count > 0)
        {
            var errorMsg = string.Join("; ", compileErrors);
            // Add hints for common type-related errors in statement lambdas
            if (ContainsTypeError(errorMsg))
            {
                errorMsg += GetStatementLambdaHint();
            }
            return new PipelineResult(false, null, $"Compile error: {errorMsg}", null);
        }

        // Track which expression this UDF name was created from
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

        // Extract input parameter from detection result
        var inputParameter = detection.Inputs.Count > 0 ? detection.Inputs[0] : null;
        // Map range ref placeholders back to original range refs
        if (inputParameter != null && detection.RangeRefMap.TryGetValue(inputParameter, out var originalRef))
        {
            inputParameter = originalRef;
        }

        // Cache the result
        _udfCache[cacheKey] = transpileResult.MethodName;
        _columnParamsCache[cacheKey] = columnParameters;
        _inputParamCache[cacheKey] = inputParameter;

        return new PipelineResult(true, transpileResult.MethodName, null, inputParameter, columnParameters);
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

    /// <summary>
    ///     Sanitizes a name to match what CodeEmitter.SanitizeName produces.
    /// </summary>
    private static string SanitizeName(string name) => CodeEmitter.SanitizeName(name);

    /// <summary>
    ///     Checks if an error message contains type-related keywords that suggest
    ///     the user might need to use helper methods in a statement lambda.
    /// </summary>
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

    /// <summary>
    ///     Returns a hint message about available helper methods for statement lambdas.
    /// </summary>
    private static string GetStatementLambdaHint() =>
        "\n\nHint: In statement lambdas, use helper methods for type conversion:\n" +
        "  Num(x) - convert to double\n" +
        "  Str(x) - convert to string\n" +
        "  Bool(x) - convert to bool\n" +
        "  Int(x) - convert to int\n" +
        "  IsEmpty(x) - check if null/empty";
}
