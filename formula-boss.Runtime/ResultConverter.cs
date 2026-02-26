namespace FormulaBoss.Runtime;

public static class ResultConverter
{
    public static object?[,] ToResult(this ExcelValue value)
    {
        return value.RawValue switch
        {
            object?[,] array => array,
            _ => new[,] { { value.RawValue } }
        };
    }

    public static object?[,] ToResult(this IExcelRange range)
    {
        if (range is ExcelValue ev)
        {
            return ev.ToResult();
        }

        var rows = range.Rows.ToList();
        if (rows.Count == 0)
        {
            return new object?[0, 0];
        }

        var cols = rows[0].ColumnCount;
        var result = new object?[rows.Count, cols];
        for (var r = 0; r < rows.Count; r++)
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = rows[r][c].Value;
            }

        return result;
    }

    // Resolves ambiguity for types that implement both ExcelValue and IExcelRange
    public static object?[,] ToResult(this ExcelScalar value) => ((ExcelValue)value).ToResult();
    public static object?[,] ToResult(this ExcelArray value) => ((ExcelValue)value).ToResult();
    public static object?[,] ToResult(this ExcelTable value) => ((ExcelValue)value).ToResult();

    public static object?[,] ToResult(this bool value) => new object?[,] { { value } };
    public static object?[,] ToResult(this int value) => new object?[,] { { value } };
    public static object?[,] ToResult(this double value) => new object?[,] { { value } };
    public static object?[,] ToResult(this string? value) => new object?[,] { { value } };

    /// <summary>
    ///     Catch-all overload for dynamic dispatch. Routes to the correct typed overload
    ///     based on runtime type, or wraps the value in a 1×1 array as a fallback.
    /// </summary>
    public static object?[,] ToResult(this object? value)
    {
        return value switch
        {
            null => new object?[,] { { null } },
            ExcelValue ev => ev.ToResult(),
            IExcelRange range => range.ToResult(),
            object?[,] array => array,
            bool b => new object?[,] { { b } },
            string s => new object?[,] { { s } },
            _ => new object?[,] { { value } }
        };
    }
}
