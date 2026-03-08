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
    public override IExcelRange Where(Func<ExcelValue, bool> predicate) =>
        predicate(this) ? this : new ExcelArray(new object[0, 0]);

    public override IExcelRange Select(Func<ExcelValue, ExcelValue> selector)
    {
        var result = selector(this);
        return result;
    }

    public override IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector)
    {
        var results = selector(this).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override bool Any(Func<ExcelValue, bool> predicate) => predicate(this);
    public override bool All(Func<ExcelValue, bool> predicate) => predicate(this);

    public override ExcelValue First(Func<ExcelValue, bool> predicate) =>
        predicate(this) ? this : throw new InvalidOperationException("No matching element.");

    public override ExcelValue? FirstOrDefault(Func<ExcelValue, bool> predicate) =>
        predicate(this) ? this : null;

    public override int Count() => 1;
    public override ExcelScalar Sum() => new(Convert.ToDouble(_value));
    public override ExcelScalar Min() => new(Convert.ToDouble(_value));
    public override ExcelScalar Max() => new(Convert.ToDouble(_value));
    public override ExcelScalar Average() => new(Convert.ToDouble(_value));

    public override IExcelRange Map(Func<ExcelValue, ExcelValue> selector)
    {
        var result = selector(this);
        return result;
    }

    public override IExcelRange OrderBy(Func<ExcelValue, object> keySelector) => this;
    public override IExcelRange OrderByDescending(Func<ExcelValue, object> keySelector) => this;
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

    /// <inheritdoc />
    public override IEnumerator<ExcelValue> GetEnumerator()
    {
        yield return this;
    }
}
