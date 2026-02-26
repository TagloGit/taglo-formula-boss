namespace FormulaBoss.Runtime;

public class ColumnValue
{
    public object? Value { get; }

    /// <summary>
    ///     Lazy cell accessor, set by <see cref="Row" /> when positional context is available.
    ///     Invokes <see cref="RuntimeBridge.GetCell" /> to escalate to COM.
    /// </summary>
    internal Func<Cell>? CellAccessor { get; init; }

    /// <summary>
    ///     Escalates to a <see cref="Cell" /> object, providing access to formatting properties.
    ///     Requires a macro-type UDF (the transpiler detects .Cell usage and sets IsMacroType).
    /// </summary>
    public Cell Cell => CellAccessor?.Invoke()
                        ?? throw new InvalidOperationException(
                            "Cell access requires a macro-type UDF with range position context.");

    public ColumnValue(object? value)
    {
        Value = value;
    }

    private double ToDouble() => Convert.ToDouble(Value);

    // Implicit conversions
    public static implicit operator double(ColumnValue v) => v.ToDouble();
    public static implicit operator string?(ColumnValue v) => v.Value?.ToString();
    public static implicit operator bool(ColumnValue v) => Convert.ToBoolean(v.Value);

    // Comparison operators
    public static bool operator ==(ColumnValue? a, ColumnValue? b) => Equals(a?.Value, b?.Value);
    public static bool operator !=(ColumnValue? a, ColumnValue? b) => !Equals(a?.Value, b?.Value);

    public static bool operator ==(ColumnValue? a, object? b)
    {
        var aVal = a?.Value;
        var bVal = b is ColumnValue cv ? cv.Value : b;

        // Numeric type coercion: double 10.0 == int 10 should be true
        if (aVal != null && bVal != null && IsNumeric(aVal) && IsNumeric(bVal))
        {
            return Convert.ToDouble(aVal) == Convert.ToDouble(bVal);
        }

        return Equals(aVal, bVal);
    }

    private static bool IsNumeric(object? value) =>
        value is double or float or int or long or short or byte or decimal;

    public static bool operator !=(ColumnValue? a, object? b) => !(a == b);
    public static bool operator ==(object? a, ColumnValue? b) => b == a;
    public static bool operator !=(object? a, ColumnValue? b) => !(b == a);

    public static bool operator >(ColumnValue a, ColumnValue b) => a.ToDouble() > b.ToDouble();
    public static bool operator <(ColumnValue a, ColumnValue b) => a.ToDouble() < b.ToDouble();
    public static bool operator >=(ColumnValue a, ColumnValue b) => a.ToDouble() >= b.ToDouble();
    public static bool operator <=(ColumnValue a, ColumnValue b) => a.ToDouble() <= b.ToDouble();

    public static bool operator >(ColumnValue a, double b) => a.ToDouble() > b;
    public static bool operator <(ColumnValue a, double b) => a.ToDouble() < b;
    public static bool operator >=(ColumnValue a, double b) => a.ToDouble() >= b;
    public static bool operator <=(ColumnValue a, double b) => a.ToDouble() <= b;

    public static bool operator >(double a, ColumnValue b) => a > b.ToDouble();
    public static bool operator <(double a, ColumnValue b) => a < b.ToDouble();
    public static bool operator >=(double a, ColumnValue b) => a >= b.ToDouble();
    public static bool operator <=(double a, ColumnValue b) => a <= b.ToDouble();

    // Arithmetic operators
    public static ColumnValue operator +(ColumnValue a, ColumnValue b) =>
        new(a.ToDouble() + b.ToDouble());

    public static ColumnValue operator -(ColumnValue a, ColumnValue b) =>
        new(a.ToDouble() - b.ToDouble());

    public static ColumnValue operator *(ColumnValue a, ColumnValue b) =>
        new(a.ToDouble() * b.ToDouble());

    public static ColumnValue operator /(ColumnValue a, ColumnValue b) =>
        new(a.ToDouble() / b.ToDouble());

    public static ColumnValue operator +(ColumnValue a, double b) => new(a.ToDouble() + b);
    public static ColumnValue operator -(ColumnValue a, double b) => new(a.ToDouble() - b);
    public static ColumnValue operator *(ColumnValue a, double b) => new(a.ToDouble() * b);
    public static ColumnValue operator /(ColumnValue a, double b) => new(a.ToDouble() / b);

    public static ColumnValue operator +(double a, ColumnValue b) => new(a + b.ToDouble());
    public static ColumnValue operator -(double a, ColumnValue b) => new(a - b.ToDouble());
    public static ColumnValue operator *(double a, ColumnValue b) => new(a * b.ToDouble());
    public static ColumnValue operator /(double a, ColumnValue b) => new(a / b.ToDouble());

    public override bool Equals(object? obj) =>
        obj is ColumnValue other ? Equals(Value, other.Value) : Equals(Value, obj);

    public override int GetHashCode() => Value?.GetHashCode() ?? 0;

    public override string ToString() => Value?.ToString() ?? "";
}
