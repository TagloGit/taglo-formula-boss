using System.Collections;

namespace FormulaBoss.Runtime;

/// <summary>Represents a single row from an Excel range, with column access via indexer or named syntax.</summary>
public class Row : ExcelArray, IEnumerable<ColumnValue>
{
    private readonly object?[] _values;

    public Row(object?[] values, Dictionary<string, int>? columnMap, Func<int, Cell>? cellResolver = null)
        : base(To2D(values), columnMap)
    {
        _values = values;
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
            if (ColumnMap == null || !ColumnMap.TryGetValue(columnName, out var index))
            {
                throw new KeyNotFoundException($"Column '{columnName}' not found.");
            }

            return MakeColumnValue(index);
        }
    }

    /// <summary>Gets a column value by zero-based index. Negative indices count from the end.</summary>
    /// <param name="index">The column index.</param>
    public new ColumnValue this[int index]
    {
        get
        {
            var i = index < 0 ? _values.Length + index : index;
            return MakeColumnValue(i);
        }
    }

    /// <summary>Gets the number of columns in this row.</summary>
    public int ColumnCount => _values.Length;

    /// <summary>Returns the row itself wrapped in a RowCollection, preserving cell resolver context.</summary>
    public override RowCollection Rows => new(new[] { this }, ColumnMap);

    /// <summary>Returns an enumerator that iterates over the column values in this row.</summary>
    IEnumerator<ColumnValue> IEnumerable<ColumnValue>.GetEnumerator()
    {
        for (var i = 0; i < _values.Length; i++)
        {
            yield return MakeColumnValue(i);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable<ExcelValue>)this).GetEnumerator();

    /// <summary>Returns an enumerator that iterates over the column values in this row as ExcelValues.</summary>
    public override IEnumerator<ExcelValue> GetEnumerator()
    {
        for (var i = 0; i < _values.Length; i++)
        {
            yield return MakeColumnValue(i);
        }
    }

    private ColumnValue MakeColumnValue(int colIndex)
    {
        var resolver = CellResolver;
        return new ColumnValue(_values[colIndex]) { CellAccessor = resolver != null ? () => resolver(colIndex) : null };
    }

    private static object?[,] To2D(object?[] values)
    {
        var result = new object?[1, values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            result[0, i] = values[i];
        }

        return result;
    }
}
