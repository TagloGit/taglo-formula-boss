namespace FormulaBoss.Runtime;

public class ExcelArray : ExcelValue, IExcelRange
{
    private readonly Dictionary<string, int>? _columnMap;
    private readonly object?[,] _data;
    private readonly RangeOrigin? _origin;

    public ExcelArray(object?[,] data, Dictionary<string, int>? columnMap = null,
        RangeOrigin? origin = null)
    {
        _data = data;
        _columnMap = columnMap;
        _origin = origin;
    }

    public override object RawValue => _data;

    public int RowCount => _data.GetLength(0);
    public int ColCount => _data.GetLength(1);

    public override RowCollection Rows
    {
        get
        {
            var rowCount = _data.GetLength(0);
            var cols = _data.GetLength(1);
            var rows = new List<Row>(rowCount);
            for (var r = 0; r < rowCount; r++)
            {
                var values = new object?[cols];
                for (var c = 0; c < cols; c++)
                {
                    values[c] = _data[r, c];
                }

                var rowIdx = r;
                Func<int, Cell>? cellResolver = _origin != null && RuntimeBridge.GetCell != null
                    ? colIdx => RuntimeBridge.GetCell(_origin.SheetName,
                        _origin.TopRow + rowIdx, _origin.LeftCol + colIdx)
                    : null;
                rows.Add(new Row(values, _columnMap, cellResolver));
            }

            return new RowCollection(rows, _columnMap);
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

            var rows = _data.GetLength(0);
            var cols = _data.GetLength(1);
            for (var r = 0; r < rows; r++)
                for (var c = 0; c < cols; c++)
                {
                    yield return RuntimeBridge.GetCell(_origin.SheetName,
                        _origin.TopRow + r, _origin.LeftCol + c);
                }
        }
    }

    public override IExcelRange Where(Func<ExcelValue, bool> predicate)
    {
        var results = ElementWise().Where(e => predicate(e)).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override IExcelRange Select(Func<ExcelValue, ExcelValue> selector)
    {
        var results = ElementWise().Select(e => selector(e)).ToList();
        if (results.Count == 0)
        {
            return new ExcelArray(new object?[0, 1]);
        }

        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue is object?[,] arr ? arr[0, 0] : results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector)
    {
        var results = ElementWise().SelectMany(e => selector(e)).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override bool Any(Func<ExcelValue, bool> predicate) =>
        ElementWise().Any(e => predicate(e));

    public override bool All(Func<ExcelValue, bool> predicate) =>
        ElementWise().All(e => predicate(e));

    public override ExcelValue First(Func<ExcelValue, bool> predicate) =>
        ElementWise().First(e => predicate(e));

    public override ExcelValue? FirstOrDefault(Func<ExcelValue, bool> predicate) =>
        ElementWise().FirstOrDefault(e => predicate(e));

    public override int Count() => _data.GetLength(0) * _data.GetLength(1);

    public override ExcelScalar Sum() => new(AggregateNumeric(0.0, (acc, v) => acc + v));
    public override ExcelScalar Min() => new(AggregateNumeric(double.MaxValue, Math.Min));
    public override ExcelScalar Max() => new(AggregateNumeric(double.MinValue, Math.Max));

    public override ExcelScalar Average()
    {
        var count = 0;
        var sum = 0.0;
        foreach (var el in ElementWise())
        {
            sum += Convert.ToDouble(el.RawValue);
            count++;
        }

        return new ExcelScalar(count == 0 ? 0.0 : sum / count);
    }

    public override IExcelRange Map(Func<ExcelValue, ExcelValue> selector)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        var result = new object?[rows, cols];
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                var mapped = selector(new ExcelScalar(_data[r, c]));
                result[r, c] = mapped.RawValue is object?[,] arr ? arr[0, 0] : mapped.RawValue;
            }

        return new ExcelArray(result, _columnMap);
    }

    public override IExcelRange OrderBy(Func<ExcelValue, object> keySelector)
    {
        var elements = ElementWise().OrderBy(e => keySelector(e)).ToList();
        return FromElements(elements);
    }

    public override IExcelRange OrderByDescending(Func<ExcelValue, object> keySelector)
    {
        var elements = ElementWise().OrderByDescending(e => keySelector(e)).ToList();
        return FromElements(elements);
    }

    public override IExcelRange Take(int count)
    {
        var elements = ElementWise().ToList();
        var taken = count >= 0 ? elements.Take(count) : elements.TakeLast(-count);
        return FromElements(taken.ToList());
    }

    public override IExcelRange Skip(int count)
    {
        var elements = ElementWise().ToList();
        var skipped = count >= 0 ? elements.Skip(count) : elements.SkipLast(-count);
        return FromElements(skipped.ToList());
    }

    public override IExcelRange Distinct()
    {
        var seen = new HashSet<string>();
        var distinct = new List<ExcelScalar>();
        foreach (var el in ElementWise())
        {
            var key = el.RawValue?.ToString() ?? "";
            if (seen.Add(key))
            {
                distinct.Add(el);
            }
        }

        return FromElements(distinct);
    }

    public override ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func)
    {
        var acc = seed;
        foreach (var el in ElementWise())
        {
            acc = func(acc, el);
        }

        return acc;
    }

    public override IExcelRange Scan(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func)
    {
        var results = new List<ExcelValue>();
        var acc = seed;
        foreach (var el in ElementWise())
        {
            acc = func(acc, el);
            results.Add(acc);
        }

        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    // --- Element-wise operations (iterate cell-by-cell, row-major) ---

    private IEnumerable<ExcelScalar> ElementWise()
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                yield return new ExcelScalar(_data[r, c]);
            }
    }

    private double AggregateNumeric(double seed, Func<double, double, double> func)
    {
        var result = seed;
        foreach (var el in ElementWise())
        {
            result = func(result, Convert.ToDouble(el.RawValue));
        }

        return result;
    }

    private static ExcelArray FromElements(List<ExcelScalar> elements)
    {
        var array = new object?[elements.Count, 1];
        for (var i = 0; i < elements.Count; i++)
        {
            array[i, 0] = elements[i].RawValue;
        }

        return new ExcelArray(array);
    }
}
