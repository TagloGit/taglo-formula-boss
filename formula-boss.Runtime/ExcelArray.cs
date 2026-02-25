namespace FormulaBoss.Runtime;

public class ExcelArray : ExcelValue, IExcelRange
{
    private readonly object?[,] _data;
    private readonly Dictionary<string, int>? _columnMap;

    public ExcelArray(object?[,] data, Dictionary<string, int>? columnMap = null)
    {
        _data = data;
        _columnMap = columnMap;
    }

    public override object? RawValue => _data;

    public int RowCount => _data.GetLength(0);
    public int ColCount => _data.GetLength(1);

    public IEnumerable<Row> Rows
    {
        get
        {
            var rows = _data.GetLength(0);
            var cols = _data.GetLength(1);
            for (var r = 0; r < rows; r++)
            {
                var values = new object?[cols];
                for (var c = 0; c < cols; c++)
                    values[c] = _data[r, c];
                yield return new Row(values, _columnMap);
            }
        }
    }

    public IExcelRange Where(Func<Row, bool> predicate) =>
        FromRows(Rows.Where(predicate));

    public IExcelRange Select(Func<Row, ExcelValue> selector)
    {
        var results = Rows.Select(selector).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
            array[i, 0] = results[i].RawValue;
        return new ExcelArray(array);
    }

    public IExcelRange SelectMany(Func<Row, IEnumerable<ExcelValue>> selector)
    {
        var results = Rows.SelectMany(selector).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
            array[i, 0] = results[i].RawValue;
        return new ExcelArray(array);
    }

    public bool Any(Func<Row, bool> predicate) => Rows.Any(predicate);
    public bool All(Func<Row, bool> predicate) => Rows.All(predicate);

    public ExcelValue First(Func<Row, bool> predicate)
    {
        var row = Rows.First(predicate);
        return RowToValue(row);
    }

    public ExcelValue? FirstOrDefault(Func<Row, bool> predicate)
    {
        var row = Rows.FirstOrDefault(predicate);
        return row == null ? null : RowToValue(row);
    }

    public int Count() => _data.GetLength(0);

    public ExcelScalar Sum() => new(AggregateNumeric(0.0, (acc, v) => acc + v));
    public ExcelScalar Min() => new(AggregateNumeric(double.MaxValue, Math.Min));
    public ExcelScalar Max() => new(AggregateNumeric(double.MinValue, Math.Max));

    public ExcelScalar Average()
    {
        var count = 0;
        var sum = 0.0;
        foreach (var row in Rows)
        {
            for (var c = 0; c < row.ColumnCount; c++)
            {
                sum += Convert.ToDouble(row[c].Value);
                count++;
            }
        }

        return new ExcelScalar(count == 0 ? 0.0 : sum / count);
    }

    public IExcelRange Map(Func<Row, ExcelValue> selector)
    {
        var rows = Rows.ToList();
        var result = new object?[rows.Count, ColCount];
        for (var r = 0; r < rows.Count; r++)
        {
            var mapped = selector(rows[r]);
            if (mapped.RawValue is object?[,] arr)
            {
                var cols = Math.Min(arr.GetLength(1), ColCount);
                for (var c = 0; c < cols; c++)
                    result[r, c] = arr[0, c];
            }
            else
            {
                result[r, 0] = mapped.RawValue;
            }
        }

        return new ExcelArray(result, _columnMap);
    }

    public IExcelRange OrderBy(Func<Row, object> keySelector) =>
        FromRows(Rows.OrderBy(keySelector));

    public IExcelRange OrderByDescending(Func<Row, object> keySelector) =>
        FromRows(Rows.OrderByDescending(keySelector));

    public IExcelRange Take(int count)
    {
        var rows = Rows.ToList();
        var taken = count >= 0 ? rows.Take(count) : rows.TakeLast(-count);
        return FromRows(taken);
    }

    public IExcelRange Skip(int count)
    {
        var rows = Rows.ToList();
        var skipped = count >= 0 ? rows.Skip(count) : rows.SkipLast(-count);
        return FromRows(skipped);
    }

    public IExcelRange Distinct()
    {
        var seen = new HashSet<string>();
        var distinctRows = new List<Row>();
        foreach (var row in Rows)
        {
            var key = string.Join("|", Enumerable.Range(0, row.ColumnCount).Select(i => row[i].Value));
            if (seen.Add(key))
                distinctRows.Add(row);
        }

        return FromRows(distinctRows);
    }

    public ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func) =>
        Rows.Aggregate(seed, (acc, row) => func(acc, row));

    public IExcelRange Scan(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func)
    {
        var results = new List<ExcelValue>();
        var acc = seed;
        foreach (var row in Rows)
        {
            acc = func(acc, row);
            results.Add(acc);
        }

        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
            array[i, 0] = results[i].RawValue;
        return new ExcelArray(array);
    }

    private double AggregateNumeric(double seed, Func<double, double, double> func)
    {
        var result = seed;
        foreach (var row in Rows)
            for (var c = 0; c < row.ColumnCount; c++)
                result = func(result, Convert.ToDouble(row[c].Value));
        return result;
    }

    private ExcelArray FromRows(IEnumerable<Row> rows)
    {
        var list = rows.ToList();
        if (list.Count == 0)
            return new ExcelArray(new object?[0, ColCount], _columnMap);

        var cols = list[0].ColumnCount;
        var result = new object?[list.Count, cols];
        for (var r = 0; r < list.Count; r++)
            for (var c = 0; c < cols; c++)
                result[r, c] = list[r][c].Value;
        return new ExcelArray(result, _columnMap);
    }

    private static ExcelValue RowToValue(Row row)
    {
        if (row.ColumnCount == 1)
            return new ExcelScalar(row[0].Value);

        var arr = new object?[1, row.ColumnCount];
        for (var c = 0; c < row.ColumnCount; c++)
            arr[0, c] = row[c].Value;
        return new ExcelArray(arr);
    }
}
