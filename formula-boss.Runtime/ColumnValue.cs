namespace FormulaBoss.Runtime;

/// <summary>Represents a single cell value from a row, with implicit conversions and formatting access via Cell.</summary>
public class ColumnValue : ExcelScalar, IComparable<ColumnValue>
{
    public ColumnValue(object? value) : base(value)
    {
    }

    /// <summary>Gets the underlying raw value of this cell.</summary>
    public object? Value => RawValue;

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

    public int CompareTo(ColumnValue? other)
    {
        if (other is null)
        {
            return 1;
        }

        // Try numeric comparison first, fall back to string comparison
        if (Value is double or int or long or float or decimal
            && other.Value is double or int or long or float or decimal)
        {
            return ToDouble().CompareTo(other.ToDouble());
        }

        return string.Compare(Value?.ToString(), other.Value?.ToString(), StringComparison.Ordinal);
    }

    private double ToDouble() => Convert.ToDouble(Value);

    // Comparison operators (ColumnValue vs ColumnValue)
    public static bool operator >(ColumnValue a, ColumnValue b) => a.ToDouble() > b.ToDouble();
    public static bool operator <(ColumnValue a, ColumnValue b) => a.ToDouble() < b.ToDouble();
    public static bool operator >=(ColumnValue a, ColumnValue b) => a.ToDouble() >= b.ToDouble();
    public static bool operator <=(ColumnValue a, ColumnValue b) => a.ToDouble() <= b.ToDouble();

    // Double comparison operators
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

    // ExcelValue comparison operators (resolves ambiguity between ColumnValue-double and ExcelValue-ExcelValue overloads)
    public static bool operator >(ColumnValue a, ExcelValue b) => Convert.ToDouble(a.RawValue) > Convert.ToDouble(b.RawValue);
    public static bool operator <(ColumnValue a, ExcelValue b) => Convert.ToDouble(a.RawValue) < Convert.ToDouble(b.RawValue);
    public static bool operator >=(ColumnValue a, ExcelValue b) => Convert.ToDouble(a.RawValue) >= Convert.ToDouble(b.RawValue);
    public static bool operator <=(ColumnValue a, ExcelValue b) => Convert.ToDouble(a.RawValue) <= Convert.ToDouble(b.RawValue);

    public static bool operator >(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) > Convert.ToDouble(b.RawValue);
    public static bool operator <(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) < Convert.ToDouble(b.RawValue);
    public static bool operator >=(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) >= Convert.ToDouble(b.RawValue);
    public static bool operator <=(ExcelValue a, ColumnValue b) => Convert.ToDouble(a.RawValue) <= Convert.ToDouble(b.RawValue);

    // Arithmetic operators (return ColumnValue to preserve type in row expressions)
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

    public override bool Equals(object? obj) =>
        obj is ExcelValue other ? Equals(RawValue, other.RawValue) : Equals(RawValue, obj);

    public override int GetHashCode() => RawValue?.GetHashCode() ?? 0;
}
