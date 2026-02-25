namespace FormulaBoss.Runtime;

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

    public IEnumerable<Cell> Cells
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

    public IEnumerable<Row> Rows
    {
        get { yield return SingleRow; }
    }

    public IExcelRange Where(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : new ExcelArray(new object[0, 0]);

    public IExcelRange Select(Func<Row, ExcelValue> selector)
    {
        var result = selector(SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }

    public IExcelRange SelectMany(Func<Row, IEnumerable<ExcelValue>> selector)
    {
        var results = selector(SingleRow).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public bool Any(Func<Row, bool> predicate) => predicate(SingleRow);
    public bool All(Func<Row, bool> predicate) => predicate(SingleRow);

    public ExcelValue First(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : throw new InvalidOperationException("No matching element.");

    public ExcelValue? FirstOrDefault(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : null;

    public int Count() => 1;
    public ExcelScalar Sum() => new(Convert.ToDouble(_value));
    public ExcelScalar Min() => new(Convert.ToDouble(_value));
    public ExcelScalar Max() => new(Convert.ToDouble(_value));
    public ExcelScalar Average() => new(Convert.ToDouble(_value));

    public IExcelRange Map(Func<Row, ExcelValue> selector)
    {
        var result = selector(SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }

    public IExcelRange OrderBy(Func<Row, object> keySelector) => this;
    public IExcelRange OrderByDescending(Func<Row, object> keySelector) => this;
    public IExcelRange Take(int count) => count == 0 ? new ExcelArray(new object[0, 0]) : this;
    public IExcelRange Skip(int count) => count == 0 ? this : new ExcelArray(new object[0, 0]);
    public IExcelRange Distinct() => this;

    public ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func) =>
        func(seed, SingleRow);

    public IExcelRange Scan(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func)
    {
        var result = func(seed, SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }
}
