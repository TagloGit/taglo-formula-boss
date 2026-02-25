using System.Dynamic;

namespace FormulaBoss.Runtime;

public class Row : DynamicObject
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

    public ColumnValue this[int index]
    {
        get
        {
            var i = index < 0 ? _values.Length + index : index;
            return MakeColumnValue(i);
        }
    }

    public int ColumnCount => _values.Length;

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
