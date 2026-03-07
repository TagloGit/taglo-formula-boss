namespace FormulaBoss.Runtime;

/// <summary>An Excel table (named range with column headers) supporting row-wise and element-wise operations.</summary>
public class ExcelTable : ExcelArray
{
    public ExcelTable(object?[,] data, string[] headers, RangeOrigin? origin = null)
        : base(data, BuildColumnMap(headers), origin)
    {
        Headers = headers;
    }

    /// <summary>Gets the column header names for this table.</summary>
    public string[] Headers { get; }

    private static Dictionary<string, int> BuildColumnMap(string[] headers)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Length; i++)
        {
            map[headers[i]] = i;
        }

        return map;
    }
}
