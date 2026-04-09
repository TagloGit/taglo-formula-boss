namespace FormulaBoss.Runtime;

/// <summary>A 2D array of Excel values supporting element-wise operations.</summary>
public class ExcelArray : ExcelValue, IExcelRange
{
    private readonly object?[,] _data;

    public ExcelArray(object?[,] data, Dictionary<string, int>? columnMap = null,
        RangeOrigin? origin = null)
    {
        _data = data;
        ColumnMap = columnMap;
        Origin = origin;
    }

    protected Dictionary<string, int>? ColumnMap { get; }

    /// <summary>Gets the range origin for this array, if available.</summary>
    protected RangeOrigin? Origin { get; }

    /// <inheritdoc />
    public override object RawValue => _data;

    public override ColumnCollection Cols
    {
        get
        {
            var data = _data;
            var rowCount = data.GetLength(0);
            var colCount = data.GetLength(1);
            var columns = new List<Column>(colCount);
            for (var c = 0; c < colCount; c++)
            {
                var colData = new object?[rowCount, 1];
                for (var r = 0; r < rowCount; r++)
                {
                    colData[r, 0] = data[r, c];
                }

                var colOrigin = Origin != null
                    ? Origin with { LeftCol = Origin.LeftCol + c }
                    : null;

                string? name = null;
                if (ColumnMap != null)
                {
                    foreach (var kvp in ColumnMap)
                    {
                        if (kvp.Value == c)
                        {
                            name = kvp.Key;
                            break;
                        }
                    }
                }

                columns.Add(new Column(colData, name, c, colOrigin));
            }

            return new ColumnCollection(columns);
        }
    }

    /// <summary>Gets the number of rows in this array.</summary>
    public override int RowCount => _data.GetLength(0);

    /// <summary>Gets the number of columns in this array.</summary>
    public override int ColCount => _data.GetLength(1);

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
                Func<int, Cell>? cellResolver = Origin != null && RuntimeBridge.GetCell != null
                    ? colIdx => RuntimeBridge.GetCell(Origin.SheetName,
                        Origin.TopRow + rowIdx, Origin.LeftCol + colIdx)
                    : null;
                rows.Add(new Row(values, ColumnMap, cellResolver, rowIdx));
            }

            return new RowCollection(rows, ColumnMap);
        }
    }

    public override IEnumerable<Cell> Cells
    {
        get
        {
            if (Origin == null || RuntimeBridge.GetCell == null)
            {
                throw new InvalidOperationException(
                    "Cell access requires a macro-type UDF with range position context.");
            }

            var rows = _data.GetLength(0);
            var cols = _data.GetLength(1);
            for (var r = 0; r < rows; r++)
                for (var c = 0; c < cols; c++)
                {
                    yield return RuntimeBridge.GetCell(Origin.SheetName,
                        Origin.TopRow + r, Origin.LeftCol + c);
                }
        }
    }

    public override IExcelRange Where(Func<ExcelScalar, bool> predicate)
    {
        var results = ElementWise().Where(e => predicate(e)).ToList();
        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i].RawValue;
        }

        return new ExcelArray(array);
    }

    public override bool Any(Func<ExcelScalar, bool> predicate) =>
        ElementWise().Any(e => predicate(e));

    public override bool All(Func<ExcelScalar, bool> predicate) =>
        ElementWise().All(e => predicate(e));

    public override ExcelValue First(Func<ExcelScalar, bool> predicate) =>
        ElementWise().First(e => predicate(e));

    public override ExcelValue? FirstOrDefault(Func<ExcelScalar, bool> predicate) =>
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

    public override IExcelRange Map(Func<ExcelScalar, ExcelScalar> selector)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        var result = new object?[rows, cols];
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                var localR = r;
                var localC = c;
                Func<Cell>? cellAccessor = Origin != null && RuntimeBridge.GetCell != null
                    ? () => RuntimeBridge.GetCell(Origin.SheetName,
                        Origin.TopRow + localR, Origin.LeftCol + localC)
                    : null;
                var scalar = new ExcelScalar(_data[r, c]) { CellAccessor = cellAccessor };
                var mapped = selector(scalar);
                result[r, c] = mapped.RawValue is object?[,] arr ? arr[0, 0] : mapped.RawValue;
            }

        return new ExcelArray(result, ColumnMap);
    }

    public override IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        var result = new object?[rows, cols];
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                var localR = r;
                var localC = c;
                Func<Cell>? cellAccessor = Origin != null && RuntimeBridge.GetCell != null
                    ? () => RuntimeBridge.GetCell(Origin.SheetName,
                        Origin.TopRow + localR, Origin.LeftCol + localC)
                    : null;
                var scalar = new ExcelScalar(_data[r, c]) { CellAccessor = cellAccessor };
                result[r, c] = selector(scalar);
            }

        return new ExcelArray(result, ColumnMap);
    }

    public override IExcelRange OrderBy(Func<ExcelScalar, object> keySelector)
    {
        var elements = ElementWise().OrderBy(e => keySelector(e)).ToList();
        return FromElements(elements);
    }

    public override IExcelRange OrderByDescending(Func<ExcelScalar, object> keySelector)
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

    public override dynamic Aggregate(dynamic seed, Func<dynamic, dynamic, dynamic> func)
    {
        var acc = seed;
        foreach (var el in ElementWise())
        {
            acc = func(acc, el);
        }

        return acc;
    }

    public override IExcelRange Scan(dynamic seed, Func<dynamic, dynamic, dynamic> func)
    {
        var results = new List<object?>();
        var acc = seed;
        foreach (var el in ElementWise())
        {
            acc = func(acc, el);
            results.Add((object?)acc);
        }

        var array = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            array[i, 0] = results[i] is ExcelValue ev ? ev.RawValue : results[i];
        }

        return new ExcelArray(array);
    }

    public override ExcelScalar this[int row, int col]
    {
        get
        {
            if (row < 0 || row >= _data.GetLength(0) || col < 0 || col >= _data.GetLength(1))
            {
                throw new IndexOutOfRangeException(
                    $"Index [{row}, {col}] is out of range for array with dimensions [{_data.GetLength(0)}, {_data.GetLength(1)}].");
            }

            return new ExcelScalar(_data[row, col]);
        }
    }

    public override ExcelScalar this[int index]
    {
        get
        {
            var cols = _data.GetLength(1);
            var total = _data.GetLength(0) * cols;
            if (index < 0 || index >= total)
            {
                throw new IndexOutOfRangeException(
                    $"Index [{index}] is out of range for array with {total} elements.");
            }

            return new ExcelScalar(_data[index / cols, index % cols]);
        }
    }

    public override int IndexOf(ExcelValue value)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        var i = 0;
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                if (Equals(_data[r, c], value.RawValue))
                {
                    return i;
                }

                i++;
            }

        return -1;
    }

    public override int IndexOf(object? value)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        var i = 0;
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                if (Equals(_data[r, c], value))
                {
                    return i;
                }

                i++;
            }

        return -1;
    }

    /// <inheritdoc />
    public override IEnumerator<ExcelValue> GetEnumerator() => ElementWise().GetEnumerator();

    /// <inheritdoc />
    public override void ForEach(Action<ExcelScalar, int, int> action)
    {
        var rows = _data.GetLength(0);
        var cols = _data.GetLength(1);
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                action(new ExcelScalar(_data[r, c]), r, c);
            }
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
