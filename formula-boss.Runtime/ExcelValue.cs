namespace FormulaBoss.Runtime;

public abstract class ExcelValue : IExcelRange, IComparable<ExcelValue>, IComparable
{
    public int CompareTo(ExcelValue? other)
    {
        if (other is null) return 1;
        var a = RawValue;
        var b = other.RawValue;
        if (a is null && b is null) return 0;
        if (a is null) return -1;
        if (b is null) return 1;
        return Comparer<double>.Default.Compare(Convert.ToDouble(a), Convert.ToDouble(b));
    }

    public int CompareTo(object? obj) => obj is ExcelValue ev ? CompareTo(ev) : 1;

    public abstract object? RawValue { get; }

    // IExcelRange — abstract, implemented by ExcelArray and ExcelScalar
    public abstract RowCollection Rows { get; }
    public abstract IEnumerable<Cell> Cells { get; }
    public abstract IExcelRange Where(Func<ExcelValue, bool> predicate);
    public abstract IExcelRange Select(Func<ExcelValue, ExcelValue> selector);
    public abstract IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector);
    public abstract bool Any(Func<ExcelValue, bool> predicate);
    public abstract bool All(Func<ExcelValue, bool> predicate);
    public abstract ExcelValue First(Func<ExcelValue, bool> predicate);
    public abstract ExcelValue? FirstOrDefault(Func<ExcelValue, bool> predicate);
    public abstract int Count();
    public abstract ExcelScalar Sum();
    public abstract ExcelScalar Min();
    public abstract ExcelScalar Max();
    public abstract ExcelScalar Average();
    public abstract IExcelRange Map(Func<ExcelValue, ExcelValue> selector);
    public abstract IExcelRange OrderBy(Func<ExcelValue, object> keySelector);
    public abstract IExcelRange OrderByDescending(Func<ExcelValue, object> keySelector);
    public abstract IExcelRange Take(int count);
    public abstract IExcelRange Skip(int count);
    public abstract IExcelRange Distinct();
    public abstract ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);
    public abstract IExcelRange Scan(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);

    public static ExcelValue Wrap(object? value, string[]? headers = null,
        RangeOrigin? origin = null)
    {
        return value switch
        {
            object[,] array when headers != null => new ExcelTable(array, headers, origin),
            object[,] array when array.GetLength(0) == 1 && array.GetLength(1) == 1
                => new ExcelScalar(array[0, 0], origin),
            object[,] array => new ExcelArray(array, origin: origin),
            ExcelValue ev => ev,
            _ => new ExcelScalar(value, origin)
        };
    }

    // Implicit conversions for scalar usage
    public static implicit operator double(ExcelValue v) => Convert.ToDouble(v.RawValue);
    public static implicit operator string?(ExcelValue v) => v.RawValue?.ToString();
    public static implicit operator bool(ExcelValue v) => Convert.ToBoolean(v.RawValue);

    // Comparison operators
    public static bool operator ==(ExcelValue? a, ExcelValue? b)
    {
        if (ReferenceEquals(a, b))
        {
            return true;
        }

        if (a is null || b is null)
        {
            return false;
        }

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

    // Arithmetic operators — return ExcelScalar so Select(x => x * 2) works
    public static ExcelScalar operator +(ExcelValue a, ExcelValue b) =>
        new(Convert.ToDouble(a.RawValue) + Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator -(ExcelValue a, ExcelValue b) =>
        new(Convert.ToDouble(a.RawValue) - Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator *(ExcelValue a, ExcelValue b) =>
        new(Convert.ToDouble(a.RawValue) * Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator /(ExcelValue a, ExcelValue b) =>
        new(Convert.ToDouble(a.RawValue) / Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator +(ExcelValue a, double b) =>
        new(Convert.ToDouble(a.RawValue) + b);

    public static ExcelScalar operator -(ExcelValue a, double b) =>
        new(Convert.ToDouble(a.RawValue) - b);

    public static ExcelScalar operator *(ExcelValue a, double b) =>
        new(Convert.ToDouble(a.RawValue) * b);

    public static ExcelScalar operator /(ExcelValue a, double b) =>
        new(Convert.ToDouble(a.RawValue) / b);

    public static ExcelScalar operator +(double a, ExcelValue b) =>
        new(a + Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator -(double a, ExcelValue b) =>
        new(a - Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator *(double a, ExcelValue b) =>
        new(a * Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator /(double a, ExcelValue b) =>
        new(a / Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator +(ExcelValue a, int b) =>
        new(Convert.ToDouble(a.RawValue) + b);

    public static ExcelScalar operator -(ExcelValue a, int b) =>
        new(Convert.ToDouble(a.RawValue) - b);

    public static ExcelScalar operator *(ExcelValue a, int b) =>
        new(Convert.ToDouble(a.RawValue) * b);

    public static ExcelScalar operator /(ExcelValue a, int b) =>
        new(Convert.ToDouble(a.RawValue) / b);

    public static ExcelScalar operator +(int a, ExcelValue b) =>
        new(a + Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator -(int a, ExcelValue b) =>
        new(a - Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator *(int a, ExcelValue b) =>
        new(a * Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator /(int a, ExcelValue b) =>
        new(a / Convert.ToDouble(b.RawValue));

    public static ExcelScalar operator -(ExcelValue a) =>
        new(-Convert.ToDouble(a.RawValue));

    // Int comparison operators (so scalar > 10 works without casting)
    public static bool operator >(ExcelValue a, int b) => Convert.ToDouble(a.RawValue) > b;
    public static bool operator <(ExcelValue a, int b) => Convert.ToDouble(a.RawValue) < b;
    public static bool operator >=(ExcelValue a, int b) => Convert.ToDouble(a.RawValue) >= b;
    public static bool operator <=(ExcelValue a, int b) => Convert.ToDouble(a.RawValue) <= b;

    public static bool operator >(int a, ExcelValue b) => a > Convert.ToDouble(b.RawValue);
    public static bool operator <(int a, ExcelValue b) => a < Convert.ToDouble(b.RawValue);
    public static bool operator >=(int a, ExcelValue b) => a >= Convert.ToDouble(b.RawValue);
    public static bool operator <=(int a, ExcelValue b) => a <= Convert.ToDouble(b.RawValue);

    public override bool Equals(object? obj) =>
        obj is ExcelValue other ? Equals(RawValue, other.RawValue) : Equals(RawValue, obj);

    public override int GetHashCode() => RawValue?.GetHashCode() ?? 0;

    public override string ToString() => RawValue?.ToString() ?? "";
}
