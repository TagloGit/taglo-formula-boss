using System.Collections;

namespace FormulaBoss.Runtime;

/// <summary>
///     Collection of <see cref="Column" /> objects with instance methods accepting <c>Func&lt;dynamic, ...&gt;</c>.
///     Symmetric with <see cref="RowCollection" /> for column-wise iteration.
/// </summary>
[SyntheticCollection(ElementType = typeof(Column))]
public class ColumnCollection : IEnumerable<Column>
{
    private readonly List<Column> _columns;

    public ColumnCollection(IEnumerable<Column> columns)
    {
        _columns = columns is List<Column> list ? list : columns.ToList();
    }

    [SyntheticExclude]
    public IEnumerator<Column> GetEnumerator() => _columns.GetEnumerator();

    [SyntheticExclude]
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Returns the number of columns.</summary>
    [SyntheticMember]
    public int Count() => _columns.Count;

    /// <summary>Filters columns to those matching the predicate.</summary>
    /// <param name="predicate">A function that receives a column (dynamic) and returns true to keep it.</param>
    [SyntheticMember]
    public ColumnCollection Where(Func<dynamic, bool> predicate) =>
        new(_columns.Where(c => predicate(c)));

    /// <summary>Projects each column into a new value.</summary>
    /// <param name="selector">A function that transforms each column.</param>
    [SyntheticMember]
    public IExcelRange Select(Func<dynamic, object> selector)
    {
        var rawResults = _columns.Select(c => selector(c)).ToList();
        if (rawResults.Count == 0)
        {
            return new ExcelArray(new object?[0, 1]);
        }

        var results = rawResults.Select(CoerceToExcelValue).ToList();
        var result = new object?[results.Count, 1];
        for (var i = 0; i < results.Count; i++)
        {
            result[i, 0] = results[i].RawValue is object?[,] arr ? arr[0, 0] : results[i].RawValue;
        }

        return new ExcelArray(result);
    }

    /// <summary>Returns the first column matching the predicate, or throws if none found.</summary>
    /// <param name="predicate">A function to test each column.</param>
    [SyntheticMember]
    public Column First(Func<dynamic, bool> predicate) =>
        _columns.First(c => predicate(c));

    /// <summary>Returns the first column matching the predicate, or null if none found.</summary>
    /// <param name="predicate">A function to test each column.</param>
    [SyntheticMember]
    public Column? FirstOrDefault(Func<dynamic, bool> predicate) =>
        _columns.FirstOrDefault(c => predicate(c));

    /// <summary>Sorts columns in ascending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each column.</param>
    [SyntheticMember]
    public ColumnCollection OrderBy(Func<dynamic, object> keySelector) =>
        new(_columns.OrderBy(c => keySelector(c)));

    /// <summary>Sorts columns in descending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each column.</param>
    [SyntheticMember]
    public ColumnCollection OrderByDescending(Func<dynamic, object> keySelector) =>
        new(_columns.OrderByDescending(c => keySelector(c)));

    /// <summary>Returns the first <paramref name="count" /> columns.</summary>
    /// <param name="count">Number of columns to take.</param>
    [SyntheticMember]
    public ColumnCollection Take(int count) =>
        new(count >= 0 ? _columns.Take(count) : _columns.TakeLast(-count));

    /// <summary>Skips the first <paramref name="count" /> columns.</summary>
    /// <param name="count">Number of columns to skip.</param>
    [SyntheticMember]
    public ColumnCollection Skip(int count) =>
        new(count >= 0 ? _columns.Skip(count) : _columns.SkipLast(-count));

    private static ExcelValue CoerceToExcelValue(object? value)
    {
        return value switch
        {
            ExcelValue ev => ev,
            _ => new ExcelScalar(value)
        };
    }
}
