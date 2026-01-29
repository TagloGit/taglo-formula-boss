using System.Runtime.InteropServices;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Manages Excel application lifecycle for integration tests.
///     Implements IDisposable to clean up Excel COM objects.
///     Use as a class fixture with xUnit: IClassFixture&lt;ExcelTestFixture&gt;
/// </summary>
public sealed class ExcelTestFixture : IDisposable
{
    private readonly dynamic _excel;
    private readonly dynamic _workbook;
    private readonly dynamic _worksheet;
    private bool _disposed;

    private int _rangeCounter;

    public ExcelTestFixture()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel is not installed or not registered");

        _excel = Activator.CreateInstance(excelType)
                 ?? throw new InvalidOperationException("Failed to create Excel Application instance");

        _excel.Visible = false;
        _excel.DisplayAlerts = false;

        _workbook = _excel.Workbooks.Add();
        _worksheet = _workbook.Worksheets[1];
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        try
        {
            _workbook.Close(false);
            if (_workbook != null)
            {
                Marshal.ReleaseComObject(_workbook);
            }

            _excel.Quit();
            if (_excel != null)
            {
                Marshal.ReleaseComObject(_excel);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    /// <summary>
    ///     Creates a range with the specified data and returns it.
    /// </summary>
    /// <param name="data">2D array of values to populate the range</param>
    /// <param name="startCell">Starting cell address (default: A1)</param>
    /// <returns>The populated Excel Range</returns>
    public dynamic CreateRange(object[,] data, string startCell = "A1")
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);

        // Calculate end cell
        var startCol = startCell[0];
        var startRow = int.Parse(startCell[1..]);
        var endCol = (char)(startCol + cols - 1);
        var endRow = startRow + rows - 1;
        var endCell = $"{endCol}{endRow}";

        var range = _worksheet.Range[$"{startCell}:{endCell}"];
        range.Value = data;
        return range;
    }

    /// <summary>
    ///     Creates a range with the specified data at a unique location (to avoid conflicts between tests).
    /// </summary>
    public dynamic CreateUniqueRange(object[,] data)
    {
        // Use a simple counter to create ranges in different columns
        var col = (char)('A' + (_rangeCounter++ % 26));
        var row = 1 + (_rangeCounter / 26 * 100); // Space out rows
        return CreateRange(data, $"{col}{row}");
    }

    /// <summary>
    ///     Sets the interior color of cells in a range.
    /// </summary>
    /// <param name="range">The range to color</param>
    /// <param name="colorIndex">Excel color index (e.g., 6 = Yellow, 3 = Red)</param>
    public void SetCellColor(dynamic range, int colorIndex) => range.Interior.ColorIndex = colorIndex;

    /// <summary>
    ///     Sets the interior color of a specific cell within a range.
    /// </summary>
    /// <param name="range">The parent range</param>
    /// <param name="row">1-based row within the range</param>
    /// <param name="col">1-based column within the range</param>
    /// <param name="colorIndex">Excel color index</param>
    public void SetCellColor(dynamic range, int row, int col, int colorIndex) =>
        range.Cells[row, col].Interior.ColorIndex = colorIndex;

    /// <summary>
    ///     Sets font bold for a cell.
    /// </summary>
    public void SetCellBold(dynamic range, int row, int col, bool bold) => range.Cells[row, col].Font.Bold = bold;
}
