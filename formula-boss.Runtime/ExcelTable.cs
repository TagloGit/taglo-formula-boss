namespace FormulaBoss.Runtime;

/// <summary>An Excel table (named range with column headers) supporting row-wise and element-wise operations.</summary>
public class ExcelTable : ExcelArray
{
    private readonly Dictionary<string, int> _headerMap;

    public ExcelTable(object?[,] data, string[] headers, RangeOrigin? origin = null)
        : base(data, BuildColumnMap(headers), origin)
    {
        Headers = headers;
        _headerMap = BuildColumnMap(headers);
    }

    /// <summary>Gets the column header names for this table.</summary>
    public string[] Headers { get; }

    /// <summary>Gets a <see cref="Column" /> by header name. Each element is a single-cell Row.</summary>
    /// <param name="columnName">The column header name (case-insensitive).</param>
    public Column this[string columnName]
    {
        get
        {
            if (!_headerMap.TryGetValue(columnName, out var colIndex))
            {
                throw new KeyNotFoundException($"Column '{columnName}' not found in table.");
            }

            return BuildColumn(columnName, colIndex);
        }
    }

    /// <summary>
    ///     Looks up a value in one column and returns the corresponding value from another column.
    ///     Similar to Excel's XLOOKUP function.
    /// </summary>
    /// <param name="lookupValue">The value to search for.</param>
    /// <param name="lookupColumn">The column to search in.</param>
    /// <param name="returnColumn">The column to return the value from.</param>
    /// <param name="ifNotFound">Value to return if no match is found (default: null).</param>
    /// <returns>The matching value from the return column, or <paramref name="ifNotFound" /> if not found.</returns>
    public object? Lookup(object? lookupValue, Column lookupColumn, Column returnColumn,
        object? ifNotFound = null)
    {
        var lookupRows = lookupColumn.ToList();
        var returnRows = returnColumn.ToList();
        var count = Math.Min(lookupRows.Count, returnRows.Count);

        for (var i = 0; i < count; i++)
        {
            if (Equals(lookupRows[i][0].Value, lookupValue))
            {
                return returnRows[i][0].Value;
            }
        }

        return ifNotFound;
    }

    private Column BuildColumn(string name, int colIndex)
    {
        var data = (object?[,])RawValue;
        var rowCount = data.GetLength(0);
        var singleColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            [name] = 0
        };
        var rows = new List<Row>(rowCount);

        for (var r = 0; r < rowCount; r++)
        {
            var values = new object?[] { data[r, colIndex] };
            rows.Add(new Row(values, singleColMap));
        }

        return new Column(name, colIndex, rows, singleColMap);
    }

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
