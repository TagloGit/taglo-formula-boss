namespace FormulaBoss.Runtime;

public class ColumnValue
{
    public object? Value { get; }

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
        return Equals(aVal, bVal);
    }

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
