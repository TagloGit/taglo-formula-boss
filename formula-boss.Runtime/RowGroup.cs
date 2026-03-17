namespace FormulaBoss.Runtime;

/// <summary>
///     A group of <see cref="Row" /> objects sharing a common key.
///     Extends <see cref="RowCollection" /> so all row-wise operations
///     (Where, Select, OrderBy, Count, ToRange, etc.) work within each group.
/// </summary>
[SyntheticCollection(ElementType = typeof(Row))]
public class RowGroup : RowCollection
{
    public RowGroup(object? key, IEnumerable<Row> rows, Dictionary<string, int>? columnMap = null)
        : base(rows, columnMap)
    {
        Key = key is ExcelValue ev ? ev.RawValue : key;
    }

    /// <summary>Gets the grouping key shared by all rows in this group.</summary>
    [SyntheticMember]
    public object? Key { get; }
}
