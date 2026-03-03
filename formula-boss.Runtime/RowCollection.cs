using System.Collections;

namespace FormulaBoss.Runtime;

/// <summary>
///     Collection of <see cref="Row" /> objects with instance methods accepting <c>Func&lt;dynamic, ...&gt;</c>.
///     Instance methods bypass the CS1977 limitation that prevents passing lambdas with dynamic
///     parameters to extension methods. The lambda parameter <c>r</c> is typed as <c>dynamic</c>,
///     enabling <c>r.ColumnName</c> and <c>r["Column Name"]</c> syntax via <see cref="Row" />'s
///     <see cref="System.Dynamic.DynamicObject" /> implementation.
/// </summary>
public class RowCollection : IEnumerable<Row>
{
    private readonly Dictionary<string, int>? _columnMap;
    private readonly List<Row> _rows;

    public RowCollection(IEnumerable<Row> rows, Dictionary<string, int>? columnMap = null)
    {
        _rows = rows is List<Row> list ? list : rows.ToList();
        _columnMap = columnMap;
    }

    public IEnumerator<Row> GetEnumerator() => _rows.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public RowCollection Where(Func<dynamic, bool> predicate) =>
        new(_rows.Where(r => predicate(r)), _columnMap);

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

    public bool Any(Func<dynamic, bool> predicate) => _rows.Any(r => predicate(r));
    public bool All(Func<dynamic, bool> predicate) => _rows.All(r => predicate(r));

    public Row First(Func<dynamic, bool> predicate) =>
        _rows.First(r => predicate(r));

    public Row? FirstOrDefault(Func<dynamic, bool> predicate) =>
        _rows.FirstOrDefault(r => predicate(r));

    public RowCollection OrderBy(Func<dynamic, object> keySelector) =>
        new(_rows.OrderBy(r => keySelector(r)), _columnMap);

    public RowCollection OrderByDescending(Func<dynamic, object> keySelector) =>
        new(_rows.OrderByDescending(r => keySelector(r)), _columnMap);

    public int Count() => _rows.Count;

    public RowCollection Take(int count) =>
        new(count >= 0 ? _rows.Take(count) : _rows.TakeLast(-count), _columnMap);

    public RowCollection Skip(int count) =>
        new(count >= 0 ? _rows.Skip(count) : _rows.SkipLast(-count), _columnMap);

    public RowCollection Distinct()
    {
        var seen = new HashSet<string>();
        var distinct = new List<Row>();
        foreach (var row in _rows)
        {
            var key = string.Join("|", Enumerable.Range(0, row.ColumnCount).Select(i => row[i].Value));
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
                result[r, c] = _rows[r][c].Value;
            }

        return new ExcelArray(result, _columnMap);
    }

    private static ExcelValue CoerceToExcelValue(object? value)
    {
        return value switch
        {
            ExcelValue ev => ev,
            ColumnValue cv => new ExcelScalar(cv.Value),
            Row row => row, // Row has implicit conversion to ExcelValue
            _ => new ExcelScalar(value)
        };
    }
}
