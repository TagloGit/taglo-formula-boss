namespace FormulaBoss.Runtime;

/// <summary>A single Excel value (one cell or a computed scalar).</summary>
public class ExcelScalar : ExcelValue, IExcelRange
{
    private readonly RangeOrigin? _origin;
    private readonly object? _value;

    public ExcelScalar(object? value, RangeOrigin? origin = null)
    {
        _value = value;
        _origin = origin;
    }

    public override object? RawValue => _value;

    // Implicit conversions from primitives
    public static implicit operator ExcelScalar(double value) => new(value);
    public static implicit operator ExcelScalar(int value) => new((double)value);
    public static implicit operator ExcelScalar(string value) => new(value);
    public static implicit operator ExcelScalar(bool value) => new(value);

    /// <summary>
    ///     Lazy cell accessor, set by <see cref="Row" /> when positional context is available.
    ///     Invokes <see cref="RuntimeBridge.GetCell" /> to escalate to COM.
    /// </summary>
    internal Func<Cell>? CellAccessor { get; init; }

    /// <summary>
    ///     Escalates to a <see cref="Cell" /> object, providing access to formatting properties.
    ///     Requires a macro-type UDF (the transpiler detects .Cell usage and sets IsMacroType).
    /// </summary>
    public Cell Cell
    {
        get
        {
            if (CellAccessor != null)
            {
                return CellAccessor.Invoke();
            }

            if (_origin != null && RuntimeBridge.GetCell != null)
            {
                return RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow, _origin.LeftCol);
            }

            throw new InvalidOperationException(
                "Cell access requires a macro-type UDF with range position context.");
        }
    }

    private Row SingleRow
    {
        get
        {
            Func<int, Cell>? cellResolver = _origin != null && RuntimeBridge.GetCell != null
                ? _ => RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow, _origin.LeftCol)
                : null;
            return new Row(new[] { _value }, null, cellResolver);
        }
    }

    public override ColumnCollection Cols =>
        new(new[] { new Column(new[,] { { _value } }, null, 0, _origin) });

    public override int RowCount => 1;
    public override int ColCount => 1;

    public override RowCollection Rows => new(new[] { SingleRow });

    public override IEnumerable<Cell> Cells
    {
        get
        {
            if (_origin == null || RuntimeBridge.GetCell == null)
            {
                throw new InvalidOperationException(
                    "Cell access requires a macro-type UDF with range position context.");
            }

            yield return RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow, _origin.LeftCol);
        }
    }

    // Element-wise: a scalar is a single element
    public override IExcelRange Where(Func<ExcelScalar, bool> predicate) =>
        predicate(this) ? this : new ExcelArray(new object[0, 0]);

    public override bool Any(Func<ExcelScalar, bool> predicate) => predicate(this);
    public override bool All(Func<ExcelScalar, bool> predicate) => predicate(this);

    public override ExcelValue First(Func<ExcelScalar, bool> predicate) =>
        predicate(this) ? this : throw new InvalidOperationException("No matching element.");

    public override ExcelValue? FirstOrDefault(Func<ExcelScalar, bool> predicate) =>
        predicate(this) ? this : null;

    public override int Count() => 1;
    public override ExcelScalar Sum() => new(Convert.ToDouble(_value));
    public override ExcelScalar Min() => new(Convert.ToDouble(_value));
    public override ExcelScalar Max() => new(Convert.ToDouble(_value));
    public override ExcelScalar Average() => new(Convert.ToDouble(_value));

    public override IExcelRange Map(Func<ExcelScalar, ExcelScalar> selector)
    {
        var result = selector(this);
        return result;
    }

    public override IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector)
    {
        var result = selector(this);
        return result is ExcelValue ev ? (IExcelRange)ev : new ExcelScalar(result);
    }

    public override IExcelRange OrderBy(Func<ExcelScalar, object> keySelector) => this;
    public override IExcelRange OrderByDescending(Func<ExcelScalar, object> keySelector) => this;
    public override IExcelRange Take(int count) => count == 0 ? new ExcelArray(new object[0, 0]) : this;
    public override IExcelRange Skip(int count) => count == 0 ? this : new ExcelArray(new object[0, 0]);
    public override IExcelRange Distinct() => this;

    public override dynamic Aggregate(dynamic seed, Func<dynamic, dynamic, dynamic> func) =>
        func(seed, this);

    public override IExcelRange Scan(dynamic seed, Func<dynamic, dynamic, dynamic> func)
    {
        var result = func(seed, this);
        if (result is IExcelRange range)
        {
            return range;
        }

        return new ExcelScalar(result is ExcelValue ev ? ev.RawValue : result);
    }

    public override ExcelScalar this[int row, int col]
    {
        get
        {
            if (row != 0 || col != 0)
            {
                throw new IndexOutOfRangeException($"ExcelScalar only supports index [0, 0], got [{row}, {col}].");
            }

            return this;
        }
    }

    public override ExcelScalar this[int index]
    {
        get
        {
            if (index != 0)
            {
                throw new IndexOutOfRangeException($"ExcelScalar only supports index [0], got [{index}].");
            }

            return this;
        }
    }

    public override int IndexOf(ExcelValue value) => Equals(RawValue, value.RawValue) ? 0 : -1;

    public override int IndexOf(object? value) => Equals(RawValue, value) ? 0 : -1;

    /// <inheritdoc />
    public override void ForEach(Action<ExcelScalar, int, int> action) => action(this, 0, 0);

    /// <inheritdoc />
    public override IEnumerator<ExcelValue> GetEnumerator()
    {
        yield return this;
    }
}
