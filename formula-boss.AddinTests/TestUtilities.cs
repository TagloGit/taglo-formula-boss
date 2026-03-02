using System.Diagnostics;
using System.Runtime.InteropServices;

using Xunit.Abstractions;

namespace FormulaBoss.AddinTests;

/// <summary>
///     Shared helpers for add-in integration tests.
/// </summary>
public static class TestUtilities
{
    /// <summary>
    ///     Enters a backtick formula into a cell as text.
    ///     Prefixes with apostrophe so Excel stores the =... literally as text,
    ///     which is how users actually enter backtick formulas.
    /// </summary>
    public static void EnterBacktickFormula(dynamic ws, string cellAddress, string formula)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            cell.Value = "'" + formula;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Polls a cell until the interceptor has rewritten it from text to a UDF call
    ///     and the UDF has returned a result, or until timeout.
    /// </summary>
    public static object? WaitForResult(dynamic ws, string cellAddress, ITestOutputHelper? output = null,
        int timeoutMs = 15000, int pollIntervalMs = 250)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            var sw = Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var formula = cell.Formula2 as string;
                    object? value = cell.Value;

                    // The interceptor rewrites the cell from text (`expr`) to a real formula (=UDF_NAME(...))
                    // Once rewritten, Formula2 starts with = but contains no backticks
                    if (formula != null && formula.StartsWith('=') && !formula.Contains('`'))
                    {
                        // Cell has been rewritten — check if the UDF has returned a value
                        if (value != null && value is not string)
                        {
                            return value;
                        }

                        // String result that isn't an error
                        if (value is string strVal && !strVal.StartsWith('#'))
                        {
                            return value;
                        }
                    }
                }
                catch
                {
                    // COM call might fail during transition
                }

                Thread.Sleep(pollIntervalMs);
            }

            // Timeout — log diagnostic info and return whatever we have
            try
            {
                var finalFormula = cell.Formula2 as string;
                var finalValue = cell.Value;
                output?.WriteLine($"TIMEOUT waiting for {cellAddress}. Formula2={finalFormula}, Value={finalValue}");
                return finalValue;
            }
            catch
            {
                return null;
            }
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Deletes a worksheet, ignoring errors during cleanup.
    /// </summary>
    public static void CleanupWorksheet(dynamic ws)
    {
        try
        {
            ws.Delete();
            Marshal.ReleaseComObject(ws);
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    /// <summary>
    ///     Gets a cell value and releases the Range COM object.
    /// </summary>
    public static object? GetCellValue(dynamic ws, string cellAddress)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            return cell.Value;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Gets a cell's Formula2 and releases the Range COM object.
    /// </summary>
    public static string? GetCellFormula(dynamic ws, string cellAddress)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            return cell.Formula2 as string;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Sets a cell value and releases the Range COM object.
    /// </summary>
    public static void SetCellValue(dynamic ws, string cellAddress, object value)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            cell.Value = value;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Sets the interior color of a cell and releases COM objects.
    /// </summary>
    public static void SetCellColor(dynamic ws, string cellAddress, int colorIndex)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            var interior = cell.Interior;
            try
            {
                interior.ColorIndex = colorIndex;
            }
            finally
            {
                Marshal.ReleaseComObject(interior);
            }
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }

    /// <summary>
    ///     Gets a cell's comment text, or null if no comment. Releases COM objects.
    /// </summary>
    public static string? GetCellComment(dynamic ws, string cellAddress)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            var comment = cell.Comment;
            if (comment == null)
            {
                return null;
            }

            try
            {
                return comment.Text() as string;
            }
            finally
            {
                Marshal.ReleaseComObject(comment);
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }
}
