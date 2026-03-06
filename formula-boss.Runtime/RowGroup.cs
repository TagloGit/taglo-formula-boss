namespace FormulaBoss.Runtime;

/// <summary>
///     A group of <see cref="Row" /> objects sharing a common key.
///     Extends <see cref="RowCollection" /> so all row-wise operations
///     (Where, Select, OrderBy, Count, ToRange, etc.) work within each group.
/// </summary>
public class RowGroup : RowCollection
{
    public RowGroup(object? key, IEnumerable<Row> rows, Dictionary<string, int>? columnMap = null)
        : base(rows, columnMap)
    {
        Key = key is ColumnValue cv ? cv.Value : key;
    }

    public object? Key { get; }
}
