namespace FormulaBoss.Transpilation;

/// <summary>
/// Result of transpiling a DSL expression to C#.
/// </summary>
/// <param name="SourceCode">The generated C# source code for the complete UDF class.</param>
/// <param name="MethodName">The generated method name (e.g., __udf_abc123).</param>
/// <param name="RequiresObjectModel">True if the expression requires Excel object model access.</param>
/// <param name="OriginalExpression">The original DSL expression that was transpiled.</param>
/// <param name="UsedColumnBindings">
/// LET variable names that were resolved as column bindings (e.g., ["price", "qty"]).
/// Used to generate UDF parameters for dynamic column name resolution.
/// </param>
public record TranspileResult(
    string SourceCode,
    string MethodName,
    bool RequiresObjectModel,
    string OriginalExpression,
    IReadOnlyList<string>? UsedColumnBindings = null);
