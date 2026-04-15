namespace FormulaBoss.Transpilation;

/// <summary>
///     Result of transpiling a DSL expression to C#.
/// </summary>
/// <param name="SourceCode">The generated C# source code for the complete UDF class.</param>
/// <param name="MethodName">The generated method name (e.g., {Prefix}abc123).</param>
/// <param name="RequiresObjectModel">True if the expression requires Excel object model access.</param>
/// <param name="OriginalExpression">The original DSL expression that was transpiled.</param>
public record TranspileResult(
    string SourceCode,
    string MethodName,
    bool RequiresObjectModel,
    string OriginalExpression)
{
    /// <summary>
    ///     Optional parallel debug-instrumented source code and method name. Populated by
    ///     <see cref="CodeEmitter.EmitDebug" /> when a debug variant is requested alongside the
    ///     normal emit.
    /// </summary>
    public TranspileResult? DebugVariant { get; init; }
}
