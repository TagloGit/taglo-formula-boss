using System.Dynamic;

namespace FormulaBoss.Runtime;

public class Row : DynamicObject
{
    private readonly object?[] _values;
    private readonly Dictionary<string, int>? _columnMap;

    public Row(object?[] values, Dictionary<string, int>? columnMap)
    {
        _values = values;
        _columnMap = columnMap;
    }

    public ColumnValue this[string columnName]
    {
        get
        {
            if (_columnMap == null || !_columnMap.TryGetValue(columnName, out var index))
                throw new KeyNotFoundException($"Column '{columnName}' not found.");
            return new ColumnValue(_values[index]);
        }
    }

    public ColumnValue this[int index]
    {
        get
        {
            var i = index < 0 ? _values.Length + index : index;
            return new ColumnValue(_values[i]);
        }
    }

    public int ColumnCount => _values.Length;

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
        if (_columnMap != null && _columnMap.TryGetValue(binder.Name, out var index))
        {
            result = new ColumnValue(_values[index]);
            return true;
        }

        result = null;
        return false;
    }

    public override IEnumerable<string> GetDynamicMemberNames() =>
        _columnMap?.Keys ?? Enumerable.Empty<string>();
}
