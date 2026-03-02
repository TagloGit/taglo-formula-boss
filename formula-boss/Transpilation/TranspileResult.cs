namespace FormulaBoss.Transpilation;

/// <summary>
/// Result of transpiling a DSL expression to C#.
/// </summary>
/// <param name="SourceCode">The generated C# source code for the complete UDF class.</param>
/// <param name="MethodName">The generated method name (e.g., __udf_abc123).</param>
/// <param name="RequiresObjectModel">True if the expression requires Excel object model access.</param>
/// <param name="OriginalExpression">The original DSL expression that was transpiled.</param>
public record TranspileResult(
    string SourceCode,
    string MethodName,
    bool RequiresObjectModel,
    string OriginalExpression);
