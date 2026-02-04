namespace FormulaBoss.Transpilation;

/// <summary>
/// Exception thrown when transpilation fails due to invalid DSL syntax or property access.
/// </summary>
public class TranspileException : Exception
{
    public TranspileException(string message) : base(message)
    {
    }

    public TranspileException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
