namespace FormulaBoss.Runtime;

/// <summary>A vertical slice of an Excel table — an N×1 array representing a single column.</summary>
public class Column : ExcelArray
{
    public Column(object?[,] data, string name, int columnIndex, RangeOrigin? origin = null)
        : base(data, origin: origin)
    {
        Name = name;
        ColumnIndex = columnIndex;
    }

    /// <summary>Gets the column header name.</summary>
    public string Name { get; }

    /// <summary>Gets the zero-based column index within the source table.</summary>
    public int ColumnIndex { get; }
}
