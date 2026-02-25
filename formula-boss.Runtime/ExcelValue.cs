namespace FormulaBoss.Runtime;

public abstract class ExcelValue
{
    public abstract object? RawValue { get; }

    public static ExcelValue Wrap(object? value, string[]? headers = null)
    {
        return value switch
        {
            object[,] array => headers != null
                ? new ExcelTable(array, headers)
                : new ExcelArray(array),
            ExcelValue ev => ev,
            _ => new ExcelScalar(value)
        };
    }

    // Implicit conversions for scalar usage
    public static implicit operator double(ExcelValue v) => Convert.ToDouble(v.RawValue);
    public static implicit operator string?(ExcelValue v) => v.RawValue?.ToString();
    public static implicit operator bool(ExcelValue v) => Convert.ToBoolean(v.RawValue);

    // Comparison operators
    public static bool operator ==(ExcelValue? a, ExcelValue? b)
    {
        if (ReferenceEquals(a, b)) return true;
        if (a is null || b is null) return false;
        return Equals(a.RawValue, b.RawValue);
    }

    public static bool operator !=(ExcelValue? a, ExcelValue? b) => !(a == b);

    public static bool operator ==(ExcelValue? a, object? b)
    {
        var bVal = b is ExcelValue ev ? ev.RawValue : b;
        return Equals(a?.RawValue, bVal);
    }

    public static bool operator !=(ExcelValue? a, object? b) => !(a == b);
    public static bool operator ==(object? a, ExcelValue? b) => b == a;
    public static bool operator !=(object? a, ExcelValue? b) => !(b == a);

    public static bool operator >(ExcelValue a, ExcelValue b) =>
        Convert.ToDouble(a.RawValue) > Convert.ToDouble(b.RawValue);

    public static bool operator <(ExcelValue a, ExcelValue b) =>
        Convert.ToDouble(a.RawValue) < Convert.ToDouble(b.RawValue);

    public static bool operator >=(ExcelValue a, ExcelValue b) =>
        Convert.ToDouble(a.RawValue) >= Convert.ToDouble(b.RawValue);

    public static bool operator <=(ExcelValue a, ExcelValue b) =>
        Convert.ToDouble(a.RawValue) <= Convert.ToDouble(b.RawValue);

    public static bool operator >(ExcelValue a, double b) => Convert.ToDouble(a.RawValue) > b;
    public static bool operator <(ExcelValue a, double b) => Convert.ToDouble(a.RawValue) < b;
    public static bool operator >=(ExcelValue a, double b) => Convert.ToDouble(a.RawValue) >= b;
    public static bool operator <=(ExcelValue a, double b) => Convert.ToDouble(a.RawValue) <= b;

    public static bool operator >(double a, ExcelValue b) => a > Convert.ToDouble(b.RawValue);
    public static bool operator <(double a, ExcelValue b) => a < Convert.ToDouble(b.RawValue);
    public static bool operator >=(double a, ExcelValue b) => a >= Convert.ToDouble(b.RawValue);
    public static bool operator <=(double a, ExcelValue b) => a <= Convert.ToDouble(b.RawValue);

    public override bool Equals(object? obj) =>
        obj is ExcelValue other ? Equals(RawValue, other.RawValue) : Equals(RawValue, obj);

    public override int GetHashCode() => RawValue?.GetHashCode() ?? 0;

    public override string ToString() => RawValue?.ToString() ?? "";
}
