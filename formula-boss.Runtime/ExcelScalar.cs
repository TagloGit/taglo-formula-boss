namespace FormulaBoss.Runtime;

public class ExcelScalar : ExcelValue
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

    public override IEnumerable<Row> Rows
    {
        get { yield return SingleRow; }
    }

    public override IExcelRange Where(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : new ExcelArray(new object[0, 0]);

    public override IExcelRange Select(Func<Row, ExcelValue> selector)
    {
        var result = selector(SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }

    public override IExcelRange SelectMany(Func<Row, IEnumerable<ExcelValue>> selector)
    {
        var results = selector(SingleRow).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override bool Any(Func<Row, bool> predicate) => predicate(SingleRow);
    public override bool All(Func<Row, bool> predicate) => predicate(SingleRow);

    public override ExcelValue First(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : throw new InvalidOperationException("No matching element.");

    public override ExcelValue? FirstOrDefault(Func<Row, bool> predicate) =>
        predicate(SingleRow) ? this : null;

    public override int Count() => 1;
    public override ExcelScalar Sum() => new(Convert.ToDouble(_value));
    public override ExcelScalar Min() => new(Convert.ToDouble(_value));
    public override ExcelScalar Max() => new(Convert.ToDouble(_value));
    public override ExcelScalar Average() => new(Convert.ToDouble(_value));

    public override IExcelRange Map(Func<Row, ExcelValue> selector)
    {
        var result = selector(SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }

    public override IExcelRange OrderBy(Func<Row, object> keySelector) => this;
    public override IExcelRange OrderByDescending(Func<Row, object> keySelector) => this;
    public override IExcelRange Take(int count) => count == 0 ? new ExcelArray(new object[0, 0]) : this;
    public override IExcelRange Skip(int count) => count == 0 ? this : new ExcelArray(new object[0, 0]);
    public override IExcelRange Distinct() => this;

    public override ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func) =>
        func(seed, SingleRow);

    public override IExcelRange Scan(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func)
    {
        var result = func(seed, SingleRow);
        return result as IExcelRange ?? new ExcelScalar(result.RawValue);
    }
}
