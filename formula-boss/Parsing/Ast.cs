namespace FormulaBoss.Parsing;

/// <summary>
/// Base class for all AST expression nodes.
/// </summary>
public abstract record Expression;

/// <summary>
/// An identifier (variable name, property name, etc.).
/// </summary>
/// <param name="Name">The identifier name.</param>
public record IdentifierExpr(string Name) : Expression;

/// <summary>
/// An Excel range reference (e.g., A1:B10, $A$1:$B$10).
/// </summary>
/// <param name="Start">The start cell reference (e.g., "A1", "$A$1").</param>
/// <param name="End">The end cell reference (e.g., "B10", "$B$10").</param>
public record RangeRefExpr(string Start, string End) : Expression;

/// <summary>
/// A numeric literal.
/// </summary>
/// <param name="Value">The numeric value.</param>
public record NumberLiteral(double Value) : Expression;

/// <summary>
/// A string literal.
/// </summary>
/// <param name="Value">The string value.</param>
public record StringLiteral(string Value) : Expression;

/// <summary>
/// A binary expression (e.g., a + b, x == y).
/// </summary>
/// <param name="Left">The left operand.</param>
/// <param name="Operator">The operator (e.g., "+", "==", "&&").</param>
/// <param name="Right">The right operand.</param>
public record BinaryExpr(Expression Left, string Operator, Expression Right) : Expression;

/// <summary>
/// A unary expression (e.g., -x, !flag).
/// </summary>
/// <param name="Operator">The operator (e.g., "-", "!").</param>
/// <param name="Operand">The operand expression.</param>
public record UnaryExpr(string Operator, Expression Operand) : Expression;

/// <summary>
/// A member access expression (e.g., obj.property).
/// </summary>
/// <param name="Target">The expression being accessed.</param>
/// <param name="Member">The member name.</param>
/// <param name="IsEscaped">If true, the property was prefixed with @ to bypass type validation.</param>
/// <param name="IsSafeAccess">If true, the property was suffixed with ? for null-safe access.</param>
public record MemberAccess(Expression Target, string Member, bool IsEscaped = false, bool IsSafeAccess = false) : Expression;

/// <summary>
/// A method call expression (e.g., obj.method(arg1, arg2)).
/// </summary>
/// <param name="Target">The expression on which the method is called.</param>
/// <param name="Method">The method name.</param>
/// <param name="Arguments">The method arguments.</param>
public record MethodCall(Expression Target, string Method, IReadOnlyList<Expression> Arguments) : Expression;

/// <summary>
/// A lambda expression (e.g., x => x.value > 0, or (acc, x) => acc + x).
/// </summary>
/// <param name="Parameters">The parameter names (one or more).</param>
/// <param name="Body">The lambda body expression.</param>
public record LambdaExpr(IReadOnlyList<string> Parameters, Expression Body) : Expression
{
    /// <summary>
    /// Convenience constructor for single-parameter lambdas.
    /// </summary>
    public LambdaExpr(string parameter, Expression body) : this(new[] { parameter }, body) { }

    /// <summary>
    /// Gets the first (or only) parameter name for backwards compatibility.
    /// </summary>
    public string Parameter => Parameters[0];
}

/// <summary>
/// A lambda with a statement block body (e.g., c => { var x = ...; return x; }).
/// The block content is passed through verbatim to C#.
/// </summary>
/// <param name="Parameters">The parameter names (one or more).</param>
/// <param name="StatementBlock">The C# statement block including braces.</param>
/// <param name="SourcePosition">The position in source where the block starts.</param>
public record StatementLambdaExpr(
    IReadOnlyList<string> Parameters,
    string StatementBlock,
    int SourcePosition) : Expression
{
    /// <summary>
    /// Convenience constructor for single-parameter statement lambdas.
    /// </summary>
    public StatementLambdaExpr(string parameter, string block, int pos)
        : this(new[] { parameter }, block, pos) { }

    /// <summary>
    /// Gets the first (or only) parameter name for backwards compatibility.
    /// </summary>
    public string Parameter => Parameters[0];
}

/// <summary>
/// A parenthesized expression (for grouping).
/// </summary>
/// <param name="Inner">The inner expression.</param>
public record GroupingExpr(Expression Inner) : Expression;

/// <summary>
/// An index access expression (e.g., row[0], array[i]).
/// </summary>
/// <param name="Target">The expression being indexed.</param>
/// <param name="Index">The index expression.</param>
public record IndexAccess(Expression Target, Expression Index) : Expression;
