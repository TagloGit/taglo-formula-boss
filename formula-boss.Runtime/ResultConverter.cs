using System.Collections;

namespace FormulaBoss.Runtime;

public static class ResultConverter
{
    /// <summary>
    ///     Shared result conversion: dispatches any result type to an Excel-compatible return value.
    ///     Scalars return bare values; multi-cell results return object[,].
    /// </summary>
    public static object Convert(object? result)
    {
        if (result == null)
        {
            return string.Empty;
        }

        if (result is ExcelValue ev)
        {
            return ev.ToResult();
        }

        if (result is IExcelRange range)
        {
            return range.ToResult();
        }

        if (result is Cell cell)
        {
            return cell.Value ?? string.Empty;
        }

        if (result is bool or int or double or string)
        {
            return result;
        }

        if (result is GroupedRowCollection grouped)
        {
            var allRows = new List<object?[]>();
            foreach (var group in grouped)
            {
                foreach (var row in group)
                {
                    var cols = row.ColumnCount;
                    var rowData = new object?[cols + 1];
                    rowData[0] = group.Key;
                    for (var c = 0; c < cols; c++)
                    {
                        rowData[c + 1] = row[c].Value;
                    }

                    allRows.Add(rowData);
                }
            }

            if (allRows.Count == 0)
            {
                return string.Empty;
            }

            var totalCols = allRows[0].Length;
            var groupArr = new object?[allRows.Count, totalCols];
            for (var r = 0; r < allRows.Count; r++)
                for (var c = 0; c < totalCols; c++)
                {
                    groupArr[r, c] = allRows[r][c];
                }

            return groupArr;
        }

        if (result is IEnumerable<Row> rows)
        {
            var rowList = rows.ToList();
            if (rowList.Count == 0)
            {
                return string.Empty;
            }

            var cols = rowList[0].ColumnCount;
            var arr = new object?[rowList.Count, cols];
            for (var r = 0; r < rowList.Count; r++)
                for (var c = 0; c < cols; c++)
                {
                    arr[r, c] = rowList[r][c].Value;
                }

            return arr;
        }

        if (result is IEnumerable<Cell> cells)
        {
            var list = cells.ToList();
            if (list.Count == 0)
            {
                return string.Empty;
            }

            var arr = new object?[list.Count, 1];
            for (var r = 0; r < list.Count; r++)
            {
                arr[r, 0] = list[r].Value;
            }

            return arr;
        }

        if (result is IEnumerable<ColumnValue> colValues)
        {
            var list = colValues.ToList();
            if (list.Count == 0)
            {
                return string.Empty;
            }

            var arr = new object?[list.Count, 1];
            for (var r = 0; r < list.Count; r++)
            {
                arr[r, 0] = list[r].Value;
            }

            return arr;
        }

        if (result is IEnumerable enumerable and not string and not object[,])
        {
            var list = enumerable.Cast<object>().ToList();
            if (list.Count == 0)
            {
                return string.Empty;
            }

            var arr = new object?[list.Count, 1];
            for (var r = 0; r < list.Count; r++)
            {
                arr[r, 0] = list[r] is ExcelValue ev2 ? ev2.RawValue
                    : list[r] is Cell c ? c.Value
                    : list[r];
            }

            return arr;
        }

        return result;
    }

    public static object ToResult(this ExcelValue value)
    {
        return value.RawValue switch
        {
            object?[,] array => array,
            _ => value.RawValue ?? string.Empty
        };
    }

    public static object ToResult(this IExcelRange range)
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
    public static object ToResult(this ExcelScalar value) => ((ExcelValue)value).ToResult();
    public static object ToResult(this ExcelArray value) => ((ExcelValue)value).ToResult();
    public static object ToResult(this ExcelTable value) => ((ExcelValue)value).ToResult();
    public static object ToResult(this Column value) => ((ExcelValue)value).ToResult();
    public static object ToResult(this ColumnValue value) => ((ExcelValue)value).ToResult();
    public static object ToResult(this Row value) => ((ExcelValue)value).ToResult();

    public static object ToResult(this bool value) => value;
    public static object ToResult(this int value) => value;
    public static object ToResult(this double value) => value;
    public static object ToResult(this string? value) => value ?? string.Empty;
}
