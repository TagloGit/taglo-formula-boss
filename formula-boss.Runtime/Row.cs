using System.Collections;
using System.Dynamic;

namespace FormulaBoss.Runtime;

/// <summary>Represents a single row from an Excel range, with column access via indexer or dynamic member syntax.</summary>
public class Row : DynamicObject, IEnumerable<ColumnValue>
{
    private readonly Dictionary<string, int>? _columnMap;
    private readonly object?[] _values;

    public Row(object?[] values, Dictionary<string, int>? columnMap, Func<int, Cell>? cellResolver = null)
    {
        _values = values;
        _columnMap = columnMap;
        CellResolver = cellResolver;
    }

    /// <summary>
    ///     Optional cell resolver: (columnIndex) → Cell.
    ///     Set when the row originates from a range with positional context.
    /// </summary>
    internal Func<int, Cell>? CellResolver { get; }

    /// <summary>Gets a column value by header name.</summary>
    /// <param name="columnName">The column header name (case-insensitive).</param>
    public ColumnValue this[string columnName]
    {
        get
        {
            if (_columnMap == null || !_columnMap.TryGetValue(columnName, out var index))
            {
                throw new KeyNotFoundException($"Column '{columnName}' not found.");
            }

            return MakeColumnValue(index);
        }
    }

    /// <summary>Gets a column value by zero-based index. Negative indices count from the end.</summary>
    /// <param name="index">The column index.</param>
    public ColumnValue this[int index]
    {
        get
        {
            var i = index < 0 ? _values.Length + index : index;
            return MakeColumnValue(i);
        }
    }

    /// <summary>Gets the number of columns in this row.</summary>
    public int ColumnCount => _values.Length;

    /// <summary>Returns an enumerator that iterates over the column values in this row.</summary>
    public IEnumerator<ColumnValue> GetEnumerator()
    {
        for (var i = 0; i < _values.Length; i++)
        {
            yield return MakeColumnValue(i);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    ///     Converts a Row to an ExcelArray (1 row × N columns) so it can be returned
    ///     from IExcelRange.Select() and other methods expecting ExcelValue.
    /// </summary>
    public static implicit operator ExcelValue(Row row)
    {
        var result = new object?[1, row._values.Length];
        for (var i = 0; i < row._values.Length; i++)
        {
            result[0, i] = row._values[i];
        }

        return new ExcelArray(result, row._columnMap);
    }

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
        if (_columnMap != null && _columnMap.TryGetValue(binder.Name, out var index))
        {
            result = MakeColumnValue(index);
            return true;
        }

        result = null;
        return false;
    }

    public override IEnumerable<string> GetDynamicMemberNames() =>
        _columnMap?.Keys ?? Enumerable.Empty<string>();

    private ColumnValue MakeColumnValue(int colIndex)
    {
        var resolver = CellResolver;
        return new ColumnValue(_values[colIndex]) { CellAccessor = resolver != null ? () => resolver(colIndex) : null };
    }
}
