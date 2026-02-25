namespace FormulaBoss.Runtime;

public class ExcelTable : ExcelArray
{
    public ExcelTable(object?[,] data, string[] headers, RangeOrigin? origin = null)
        : base(data, BuildColumnMap(headers), origin)
    {
        Headers = headers;
    }

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
