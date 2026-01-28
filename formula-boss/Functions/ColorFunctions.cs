using ExcelDna.Integration;

namespace FormulaBoss.Functions;

/// <summary>
/// Proof-of-concept UDFs demonstrating ExcelDNA + COM interop for cell property access.
/// </summary>
public static class ColorFunctions
{
    /// <summary>
    /// Filters cells by interior color index, returning matching values as a spill array.
    /// </summary>
    /// <param name="rangeRef">Range reference (received via AllowReference).</param>
    /// <param name="colorIndex">Excel ColorIndex to match (e.g., 6 = yellow).</param>
    /// <returns>2D array of matching cell values.</returns>
    [ExcelFunction(
        Name = "FilterByColor",
        Description = "Returns values from cells matching the specified color index")]
    public static object FilterByColor(
        [ExcelArgument(AllowReference = true, Description = "Range to filter")] object rangeRef,
        [ExcelArgument(Description = "ColorIndex to match (e.g., 6 = yellow)")] int colorIndex)
    {
        if (rangeRef is not ExcelReference excelRef)
        {
            return ExcelError.ExcelErrorRef;
        }

        try
        {
            var app = (dynamic)ExcelDnaUtil.Application;
            var sheet = app.ActiveSheet;

            // Convert ExcelReference to A1-style address
            var address = GetRangeAddress(excelRef);
            var range = sheet.Range[address];

            var matchingValues = new List<object>();

            foreach (var cell in range.Cells)
            {
                var cellColorIndex = (int)cell.Interior.ColorIndex;
                if (cellColorIndex == colorIndex)
                {
                    var value = cell.Value2;
                    matchingValues.Add(value ?? string.Empty);
                }
            }

            if (matchingValues.Count == 0)
            {
                return ExcelError.ExcelErrorNA;
            }

            // Return as single-column 2D array for spill
            var result = new object[matchingValues.Count, 1];
            for (var i = 0; i < matchingValues.Count; i++)
            {
                result[i, 0] = matchingValues[i];
            }

            return result;
        }
        catch (Exception ex)
        {
            // In production, we'd log this. For now, return error with message hint.
            System.Diagnostics.Debug.WriteLine($"FilterByColor error: {ex.Message}");
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Converts an ExcelReference to an A1-style range address.
    /// </summary>
    private static string GetRangeAddress(ExcelReference excelRef)
    {
        var rowFirst = excelRef.RowFirst + 1;  // Excel is 1-based
        var rowLast = excelRef.RowLast + 1;
        var colFirst = excelRef.ColumnFirst + 1;
        var colLast = excelRef.ColumnLast + 1;

        var startCell = GetCellAddress(rowFirst, colFirst);
        var endCell = GetCellAddress(rowLast, colLast);

        return startCell == endCell ? startCell : $"{startCell}:{endCell}";
    }

    /// <summary>
    /// Converts row/column numbers to A1-style cell address.
    /// </summary>
    private static string GetCellAddress(int row, int col)
    {
        var colLetter = GetColumnLetter(col);
        return $"{colLetter}{row}";
    }

    /// <summary>
    /// Converts a 1-based column number to Excel column letters (A, B, ..., Z, AA, AB, ...).
    /// </summary>
    private static string GetColumnLetter(int col)
    {
        var result = string.Empty;
        while (col > 0)
        {
            col--;
            result = (char)('A' + (col % 26)) + result;
            col /= 26;
        }

        return result;
    }
}
