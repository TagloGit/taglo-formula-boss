namespace FormulaBoss.Runtime;

/// <summary>
///     Represents a single column from an <see cref="ExcelTable" />.
///     Each "row" is a single-cell <see cref="Row" />, so all <see cref="RowCollection" /> methods
///     (Where, Select, Aggregate, OrderBy, etc.) work for free.
/// </summary>
public class Column : RowCollection
{
    public Column(string name, int columnIndex, IEnumerable<Row> rows,
        Dictionary<string, int>? columnMap = null)
        : base(rows, columnMap)
    {
        Name = name;
        ColumnIndex = columnIndex;
    }

    /// <summary>Gets the column header name.</summary>
    public string Name { get; }

    /// <summary>Gets the zero-based column index within the source table.</summary>
    public int ColumnIndex { get; }
}
