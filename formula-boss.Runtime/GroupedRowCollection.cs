using System.Collections;

namespace FormulaBoss.Runtime;

/// <summary>
///     Collection of <see cref="RowGroup" /> objects with instance methods accepting <c>Func&lt;dynamic, ...&gt;</c>.
///     Mirrors <see cref="RowCollection" />'s pattern to bypass the CS1977 limitation.
///     The lambda parameter is typed as <c>dynamic</c> and receives a <see cref="RowGroup" />,
///     enabling <c>g.Key</c>, <c>g.Count()</c>, <c>g.Where(...)</c>, etc.
/// </summary>
public class GroupedRowCollection : IEnumerable<RowGroup>
{
    private readonly List<RowGroup> _groups;

    public GroupedRowCollection(List<RowGroup> groups)
    {
        _groups = groups;
    }

    public IEnumerator<RowGroup> GetEnumerator() => _groups.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IExcelRange Select(Func<dynamic, object> selector)
    {
        var rawResults = _groups.Select(g => selector(g)).ToList();
        if (rawResults.Count == 0)
        {
            return new ExcelArray(new object?[0, 1]);
        }

        var results = rawResults.Select(CoerceToExcelValue).ToList();

        // Multi-column detection (same as RowCollection.Select)
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

            return new ExcelArray(array);
        }

        var result = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            result[i, 0] = results[i].RawValue is object?[,] arr ? arr[0, 0] : results[i].RawValue;
        }

        return new ExcelArray(result);
    }

    public GroupedRowCollection Where(Func<dynamic, bool> predicate) =>
        new(_groups.Where(g => predicate(g)).ToList());

    public GroupedRowCollection OrderBy(Func<dynamic, object> keySelector) =>
        new(_groups.OrderBy(g => keySelector(g)).ToList());

    public GroupedRowCollection OrderByDescending(Func<dynamic, object> keySelector) =>
        new(_groups.OrderByDescending(g => keySelector(g)).ToList());

    public int Count() => _groups.Count;

    public RowGroup First() => _groups.First();

    public RowGroup First(Func<dynamic, bool> predicate) =>
        _groups.First(g => predicate(g));

    private static ExcelValue CoerceToExcelValue(object? value)
    {
        return value switch
        {
            ExcelValue ev => ev,
            ColumnValue cv => new ExcelScalar(cv.Value),
            Row row => row,
            object?[] arr => ToSingleRowArray(arr),
            _ => new ExcelScalar(value)
        };
    }

    private static ExcelArray ToSingleRowArray(object?[] values)
    {
        var result = new object?[1, values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            result[0, i] = values[i] is ColumnValue cv ? cv.Value : values[i];
        }

        return new ExcelArray(result);
    }
}
