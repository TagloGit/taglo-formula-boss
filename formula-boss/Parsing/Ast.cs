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
public record MemberAccess(Expression Target, string Member) : Expression;

/// <summary>
/// A method call expression (e.g., obj.method(arg1, arg2)).
/// </summary>
/// <param name="Target">The expression on which the method is called.</param>
/// <param name="Method">The method name.</param>
/// <param name="Arguments">The method arguments.</param>
public record MethodCall(Expression Target, string Method, IReadOnlyList<Expression> Arguments) : Expression;

/// <summary>
/// A lambda expression (e.g., x => x.value > 0).
/// </summary>
/// <param name="Parameter">The parameter name.</param>
/// <param name="Body">The lambda body expression.</param>
public record LambdaExpr(string Parameter, Expression Body) : Expression;

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
