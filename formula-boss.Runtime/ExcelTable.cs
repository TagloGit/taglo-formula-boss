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

    /// <summary>Gets a column by header name as an N×1 array.</summary>
    /// <param name="columnName">The column header name (case-insensitive).</param>
    public Column this[string columnName]
    {
        get
        {
            if (ColumnMap == null || !ColumnMap.TryGetValue(columnName, out var colIndex))
            {
                throw new KeyNotFoundException($"Column '{columnName}' not found in table.");
            }

            var data = (object?[,])RawValue;
            var rows = data.GetLength(0);
            var colData = new object?[rows, 1];
            for (var r = 0; r < rows; r++)
            {
                colData[r, 0] = data[r, colIndex];
            }

            var colOrigin = Origin != null
                ? Origin with { LeftCol = Origin.LeftCol + colIndex }
                : null;

            return new Column(colData, columnName, colIndex, colOrigin);
        }
    }

    /// <summary>Gets columns with names from the table headers.</summary>
    public override ColumnCollection Cols
    {
        get
        {
            var data = (object?[,])RawValue;
            var rowCount = data.GetLength(0);
            var columns = new List<Column>(Headers.Length);
            for (var c = 0; c < Headers.Length; c++)
            {
                var colData = new object?[rowCount, 1];
                for (var r = 0; r < rowCount; r++)
                {
                    colData[r, 0] = data[r, c];
                }

                var colOrigin = Origin != null
                    ? Origin with { LeftCol = Origin.LeftCol + c }
                    : null;

                columns.Add(new Column(colData, Headers[c], c, colOrigin));
            }

            return new ColumnCollection(columns);
        }
    }

    /// <summary>
    ///     Looks up a value in one column and returns the corresponding value from another column.
    ///     Matches the first occurrence, similar to XLOOKUP.
    /// </summary>
    /// <param name="value">The value to search for.</param>
    /// <param name="lookupColumn">The column to search in.</param>
    /// <param name="returnColumn">The column to return the value from.</param>
    /// <param name="ifNotFound">Value to return if no match is found. If null and no match, throws.</param>
    public static ExcelScalar Lookup(object value, Column lookupColumn, Column returnColumn,
        object? ifNotFound = null)
    {
        if (lookupColumn.RowCount != returnColumn.RowCount)
        {
            throw new ArgumentException(
                $"Lookup and return columns must have the same number of rows " +
                $"(lookup: {lookupColumn.RowCount}, return: {returnColumn.RowCount}).");
        }

        var searchValue = value is ExcelValue ev ? ev.RawValue : value;
        var lookupData = (object?[,])lookupColumn.RawValue;
        var returnData = (object?[,])returnColumn.RawValue;

        for (var r = 0; r < lookupColumn.RowCount; r++)
        {
            if (IsMatch(searchValue, lookupData[r, 0]))
            {
                return new ExcelScalar(returnData[r, 0]);
            }
        }

        if (ifNotFound != null)
        {
            return new ExcelScalar(ifNotFound is ExcelValue ev2 ? ev2.RawValue : ifNotFound);
        }

        throw new KeyNotFoundException(
            $"Value '{searchValue}' not found in lookup column '{lookupColumn.Name}'.");
    }

    private static bool IsMatch(object? search, object? candidate)
    {
        if (search is string s1 && candidate is string s2)
        {
            return s1.Equals(s2, StringComparison.OrdinalIgnoreCase);
        }

        return Equals(search, candidate);
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
