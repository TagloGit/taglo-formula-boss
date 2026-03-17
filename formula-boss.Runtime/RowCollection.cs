using System.Collections;

namespace FormulaBoss.Runtime;

/// <summary>
///     Collection of <see cref="Row" /> objects with instance methods accepting <c>Func&lt;dynamic, ...&gt;</c>.
///     Instance methods bypass the CS1977 limitation that prevents passing lambdas with dynamic
///     parameters to extension methods. The lambda parameter <c>r</c> is typed as <c>dynamic</c>,
///     enabling <c>r["Column Name"]</c> syntax via <see cref="Row" />'s indexer.
/// </summary>
[SyntheticCollection(ElementType = typeof(Row))]
public class RowCollection : IEnumerable<Row>
{
    private readonly Dictionary<string, int>? _columnMap;
    private readonly List<Row> _rows;

    public RowCollection(IEnumerable<Row> rows, Dictionary<string, int>? columnMap = null)
    {
        _rows = rows is List<Row> list ? list : rows.ToList();
        _columnMap = columnMap;
    }

    [SyntheticExclude]
    public IEnumerator<Row> GetEnumerator() => _rows.GetEnumerator();

    [SyntheticExclude]
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Filters rows to those matching the predicate.</summary>
    /// <param name="predicate">A function that receives a row (dynamic) and returns true to keep it.</param>
    [SyntheticMember]
    public RowCollection Where(Func<dynamic, bool> predicate) =>
        new(_rows.Where(r => predicate(r)), _columnMap);

    /// <summary>Projects each row into a new value. Use <c>r => r.Column</c> to extract a column, or <c>r => r</c> to return full rows.</summary>
    /// <param name="selector">A function that transforms each row.</param>
    /// <returns>A range containing the projected values.</returns>
    [SyntheticMember]
    public IExcelRange Select(Func<dynamic, object> selector)
    {
        var rawResults = _rows.Select(r => selector(r)).ToList();
        if (rawResults.Count == 0)
        {
            return new ExcelArray(new object?[0, 1]);
        }

        // Coerce results to ExcelValue
        var results = rawResults.Select(CoerceToExcelValue).ToList();

        // Check if results are multi-column (identity select returning full rows)
        if (results[0].RawValue is object?[,] firstArr && firstArr.GetLength(1) > 1)
        {
            var cols = firstArr.GetLength(1);
            var array = new object?[results.Count, cols];
            for (var r = 0; r < results.Count; r++)
            {
                if (results[r].RawValue is object?[,] rowArr)
                {
                    for (var c = 0; c < Math.Min(rowArr.GetLength(1), cols); c++)
                    {
                        array[r, c] = rowArr[0, c];
                    }
                }
                else
                {
                    array[r, 0] = results[r].RawValue;
                }
            }

            return new ExcelArray(array, _columnMap);
        }

        var result = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            result[i, 0] = results[i].RawValue is object?[,] arr ? arr[0, 0] : results[i].RawValue;
        }

        return new ExcelArray(result);
    }

    /// <summary>Returns true if any row matches the predicate.</summary>
    /// <param name="predicate">A function to test each row.</param>
    [SyntheticMember]
    public bool Any(Func<dynamic, bool> predicate) => _rows.Any(r => predicate(r));

    /// <summary>Returns true if all rows match the predicate.</summary>
    /// <param name="predicate">A function to test each row.</param>
    [SyntheticMember]
    public bool All(Func<dynamic, bool> predicate) => _rows.All(r => predicate(r));

    /// <summary>Returns the first row matching the predicate, or throws if none found.</summary>
    /// <param name="predicate">A function to test each row.</param>
    [SyntheticMember]
    public Row First(Func<dynamic, bool> predicate) =>
        _rows.First(r => predicate(r));

    /// <summary>Returns the first row matching the predicate, or null if none found.</summary>
    /// <param name="predicate">A function to test each row.</param>
    [SyntheticMember]
    public Row? FirstOrDefault(Func<dynamic, bool> predicate) =>
        _rows.FirstOrDefault(r => predicate(r));

    /// <summary>Sorts rows in ascending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each row.</param>
    [SyntheticMember]
    public RowCollection OrderBy(Func<dynamic, object> keySelector) =>
        new(_rows.OrderBy(r => keySelector(r)), _columnMap);

    /// <summary>Sorts rows in descending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each row.</param>
    [SyntheticMember]
    public RowCollection OrderByDescending(Func<dynamic, object> keySelector) =>
        new(_rows.OrderByDescending(r => keySelector(r)), _columnMap);

    /// <summary>Returns the number of rows.</summary>
    [SyntheticMember]
    public int Count() => _rows.Count;

    /// <summary>Returns the first <paramref name="count" /> rows. If negative, returns the last N rows.</summary>
    /// <param name="count">Number of rows to take.</param>
    [SyntheticMember]
    public RowCollection Take(int count) =>
        new(count >= 0 ? _rows.Take(count) : _rows.TakeLast(-count), _columnMap);

    /// <summary>Skips the first <paramref name="count" /> rows. If negative, skips the last N rows.</summary>
    /// <param name="count">Number of rows to skip.</param>
    [SyntheticMember]
    public RowCollection Skip(int count) =>
        new(count >= 0 ? _rows.Skip(count) : _rows.SkipLast(-count), _columnMap);

    /// <summary>Groups rows by a key, returning a <see cref="GroupedRowCollection" /> for per-group operations.</summary>
    /// <param name="keySelector">A function that extracts the grouping key from each row.</param>
    [SyntheticMember]
    public GroupedRowCollection GroupBy(Func<dynamic, object> keySelector) =>
        new(_rows.GroupBy(r => keySelector(r), r => r)
            .Select(g => new RowGroup(g.Key, g, _columnMap))
            .ToList());

    /// <summary>Applies an accumulator function over the rows, returning the final result.</summary>
    /// <param name="seed">The initial accumulator value.</param>
    /// <param name="func">A function that takes (accumulator, row) and returns the new accumulator.</param>
    [SyntheticMember]
    public dynamic Aggregate(dynamic seed, Func<dynamic, dynamic, dynamic> func)
    {
        dynamic acc = seed;
        foreach (var row in _rows)
        {
            acc = func(acc, row);
        }

        return acc;
    }

    /// <summary>Like Aggregate, but returns all intermediate accumulator values as a range.</summary>
    /// <param name="seed">The initial accumulator value.</param>
    /// <param name="func">A function that takes (accumulator, row) and returns the new accumulator.</param>
    [SyntheticMember]
    public IExcelRange Scan(dynamic seed, Func<dynamic, dynamic, dynamic> func)
    {
        var results = new List<object?>();
        dynamic acc = seed;
        foreach (var row in _rows)
        {
            acc = func(acc, row);
            results.Add((object?)acc);
        }

        var arr = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            arr[i, 0] = results[i] is ExcelValue ev ? ev.RawValue : results[i];
        }

        return new ExcelArray(arr);
    }

    /// <summary>Returns distinct rows (compared by concatenated column values).</summary>
    [SyntheticMember]
    public RowCollection Distinct()
    {
        var seen = new HashSet<string>();
        var distinct = new List<Row>();
        foreach (var row in _rows)
        {
            var key = string.Join("|", Enumerable.Range(0, row.ColumnCount).Select(i => row[i].RawValue));
            if (seen.Add(key))
            {
                distinct.Add(row);
            }
        }

        return new RowCollection(distinct, _columnMap);
    }

    /// <summary>
    ///     Convert back to an <see cref="ExcelArray" /> for further element-wise operations.
    /// </summary>
    [SyntheticMember]
    public IExcelRange ToRange()
    {
        if (_rows.Count == 0)
        {
            return new ExcelArray(new object?[0, 1], _columnMap);
        }

        var cols = _rows[0].ColumnCount;
        var result = new object?[_rows.Count, cols];
        for (var r = 0; r < _rows.Count; r++)
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = _rows[r][c].RawValue;
            }

        return new ExcelArray(result, _columnMap);
    }

    private static ExcelValue CoerceToExcelValue(object? value)
    {
        return value switch
        {
            ExcelValue ev => ev,
            _ => new ExcelScalar(value)
        };
    }
}
