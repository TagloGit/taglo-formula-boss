namespace FormulaBoss.Runtime;

/// <summary>A vertical slice of an Excel table — an N×1 array representing a single column.</summary>
public class Column : ExcelArray
{
    public Column(object?[,] data, string? name, int index, RangeOrigin? origin = null)
        : base(data, origin: origin)
    {
        Name = name;
        Index = index;
    }

    /// <summary>Gets the column header name, or null for plain arrays without headers.</summary>
    public string? Name { get; }

    /// <summary>Gets the zero-based column index within the source array or table.</summary>
    public int Index { get; }
}
