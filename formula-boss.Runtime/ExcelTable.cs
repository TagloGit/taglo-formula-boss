namespace FormulaBoss.Runtime;

public class ExcelTable : ExcelArray
{
    public string[] Headers { get; }

    public ExcelTable(object?[,] data, string[] headers)
        : base(data, BuildColumnMap(headers))
    {
        Headers = headers;
    }

    private static Dictionary<string, int> BuildColumnMap(string[] headers)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Length; i++)
            map[headers[i]] = i;
        return map;
    }
}
