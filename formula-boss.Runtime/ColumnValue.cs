namespace FormulaBoss.Runtime;

public class ColumnValue : IComparable<ColumnValue>, IComparable
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

    // Int comparison operators (so r[0] > 10 works without casting)
    public static bool operator >(ColumnValue a, int b) => a.ToDouble() > b;
    public static bool operator <(ColumnValue a, int b) => a.ToDouble() < b;
    public static bool operator >=(ColumnValue a, int b) => a.ToDouble() >= b;
    public static bool operator <=(ColumnValue a, int b) => a.ToDouble() <= b;

    public static bool operator >(int a, ColumnValue b) => a > b.ToDouble();
    public static bool operator <(int a, ColumnValue b) => a < b.ToDouble();
    public static bool operator >=(int a, ColumnValue b) => a >= b.ToDouble();
    public static bool operator <=(int a, ColumnValue b) => a <= b.ToDouble();

    // ExcelValue comparison operators (so r[0] > maxVal works when maxVal is ExcelValue)
    public static bool operator >(ColumnValue a, ExcelValue b) => a.ToDouble() > Convert.ToDouble(b.RawValue);
    public static bool operator <(ColumnValue a, ExcelValue b) => a.ToDouble() < Convert.ToDouble(b.RawValue);
    public static bool operator >=(ColumnValue a, ExcelValue b) => a.ToDouble() >= Convert.ToDouble(b.RawValue);
    public static bool operator <=(ColumnValue a, ExcelValue b) => a.ToDouble() <= Convert.ToDouble(b.RawValue);

    public static bool operator >(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) > b.ToDouble();
    public static bool operator <(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) < b.ToDouble();
    public static bool operator >=(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) >= b.ToDouble();
    public static bool operator <=(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) <= b.ToDouble();

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

    // Int arithmetic operators
    public static ColumnValue operator +(ColumnValue a, int b) => new(a.ToDouble() + b);
    public static ColumnValue operator -(ColumnValue a, int b) => new(a.ToDouble() - b);
    public static ColumnValue operator *(ColumnValue a, int b) => new(a.ToDouble() * b);
    public static ColumnValue operator /(ColumnValue a, int b) => new(a.ToDouble() / b);

    public static ColumnValue operator +(int a, ColumnValue b) => new(a + b.ToDouble());
    public static ColumnValue operator -(int a, ColumnValue b) => new(a - b.ToDouble());
    public static ColumnValue operator *(int a, ColumnValue b) => new(a * b.ToDouble());
    public static ColumnValue operator /(int a, ColumnValue b) => new(a / b.ToDouble());

    public int CompareTo(ColumnValue? other)
    {
        if (other is null) return 1;
        // Try numeric comparison first, fall back to string comparison
        if (Value is double or int or long or float or decimal
            && other.Value is double or int or long or float or decimal)
        {
            return ToDouble().CompareTo(other.ToDouble());
        }

        return string.Compare(Value?.ToString(), other.Value?.ToString(), StringComparison.Ordinal);
    }

    public int CompareTo(object? obj)
    {
        if (obj is ColumnValue cv) return CompareTo(cv);
        if (Value is double or int or long or float or decimal)
        {
            try { return Convert.ToDouble(Value).CompareTo(Convert.ToDouble(obj)); }
            catch { /* fall through to string */ }
        }

        return string.Compare(Value?.ToString(), obj?.ToString(), StringComparison.Ordinal);
    }

    public override bool Equals(object? obj) =>
        obj is ColumnValue other ? Equals(Value, other.Value) : Equals(Value, obj);

    public override int GetHashCode() => Value?.GetHashCode() ?? 0;

    public override string ToString() => Value?.ToString() ?? "";
}
