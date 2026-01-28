using System.Diagnostics;

using FormulaBoss.Compilation;
using FormulaBoss.Parsing;
using FormulaBoss.Transpilation;

namespace FormulaBoss.Interception;

/// <summary>
/// Result of processing a backtick expression through the pipeline.
/// </summary>
/// <param name="Success">Whether processing succeeded.</param>
/// <param name="UdfName">The generated UDF name (if successful).</param>
/// <param name="ErrorMessage">Error message (if failed).</param>
/// <param name="InputParameter">The input parameter name extracted from the expression.</param>
public record PipelineResult(bool Success, string? UdfName, string? ErrorMessage, string? InputParameter);

/// <summary>
/// Orchestrates the complete pipeline: parse → transpile → compile → register.
/// </summary>
public class FormulaPipeline
{
    private readonly DynamicCompiler _compiler;
    private readonly Dictionary<string, string> _udfCache = new();

    public FormulaPipeline(DynamicCompiler compiler)
    {
        _compiler = compiler;
    }

    /// <summary>
    /// Processes a DSL expression and returns the UDF name to use.
    /// </summary>
    /// <param name="expression">The DSL expression (without backticks).</param>
    /// <returns>The pipeline result.</returns>
    public PipelineResult Process(string expression)
    {
        // Check cache first
        if (_udfCache.TryGetValue(expression, out var cachedUdfName))
        {
            // Extract input parameter from expression for cache hit
            var inputParam = ExtractInputParameter(expression);
            return new PipelineResult(true, cachedUdfName, null, inputParam);
        }

        // Step 1: Lex
        var lexer = new Lexer(expression);
        var tokens = lexer.ScanTokens();

        // Check for lexer errors
        var errorToken = tokens.FirstOrDefault(t => t.Type == TokenType.Error);
        if (errorToken != null)
        {
            return new PipelineResult(false, null, $"Lexer error: {errorToken.Lexeme}", null);
        }

        // Step 2: Parse
        var parser = new Parser(tokens);
        var ast = parser.Parse();

        if (ast == null || parser.Errors.Count > 0)
        {
            var errorMsg = parser.Errors.Count > 0
                ? string.Join("; ", parser.Errors)
                : "Unknown parse error";
            return new PipelineResult(false, null, $"Parse error: {errorMsg}", null);
        }

        // Step 3: Transpile
        var transpiler = new CSharpTranspiler();
        TranspileResult transpileResult;
        try
        {
            transpileResult = transpiler.Transpile(ast, expression);
        }
        catch (Exception ex)
        {
            return new PipelineResult(false, null, $"Transpile error: {ex.Message}", null);
        }

        // Debug: Output the generated source code
        Debug.WriteLine("=== Generated UDF Source Code ===");
        Debug.WriteLine(transpileResult.SourceCode);
        Debug.WriteLine("=== End Generated Code ===");

        // Step 4: Compile and Register
        var compileErrors = _compiler.CompileAndRegister(transpileResult.SourceCode);

        if (compileErrors.Count > 0)
        {
            var errorMsg = string.Join("; ", compileErrors);
            return new PipelineResult(false, null, $"Compile error: {errorMsg}", null);
        }

        // Cache the result
        _udfCache[expression] = transpileResult.MethodName;

        // Extract input parameter from the expression
        var inputParameter = ExtractInputParameter(expression);

        return new PipelineResult(true, transpileResult.MethodName, null, inputParameter);
    }

    /// <summary>
    /// Extracts the input parameter (range reference) from a DSL expression.
    /// </summary>
    private static string ExtractInputParameter(string expression)
    {
        // The input parameter is the first identifier before .cells or .values
        // e.g., "A1:J10.cells.where(...)" → "A1:J10"
        // e.g., "data.values.where(...)" → "data"

        var dotIndex = expression.IndexOf('.');
        if (dotIndex > 0)
        {
            return expression[..dotIndex];
        }

        return expression;
    }
}
