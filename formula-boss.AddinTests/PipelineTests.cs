using System.Diagnostics;

using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.AddinTests;

/// <summary>
///     Core pipeline tests that exercise the full add-in loop with more complex expressions.
///     Covers object model path (cell colors), table expressions, and error handling.
/// </summary>
[Collection("Excel Addin")]
public class PipelineTests
{
    private readonly ExcelAddinFixture _excel;
    private readonly ITestOutputHelper _output;

    public PipelineTests(ExcelAddinFixture excel, ITestOutputHelper output)
    {
        _excel = excel;
        _output = output;
    }

    [Fact]
    public void ObjectModelPath_WhereColor_Sum()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up data with colors
            ws.Range["A1"].Value = 10.0;
            ws.Range["A2"].Value = 20.0;
            ws.Range["A3"].Value = 30.0;
            ws.Range["A4"].Value = 40.0;
            ws.Range["A5"].Value = 50.0;

            // Color cells A2 and A4 yellow (ColorIndex 6)
            ws.Range["A2"].Interior.ColorIndex = 6;
            ws.Range["A4"].Interior.ColorIndex = 6;

            // Enter expression that filters by color and sums
            EnterBacktickFormula(ws, "B1", "=`A1:A5.cells.where(c => c.color == 6).sum()`");

            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 formula: {ws.Range["B1"].Formula2}");
            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result)); // 20 + 40
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ObjectModelPath_WhereColor_Count()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            ws.Range["A1"].Value = 1.0;
            ws.Range["A2"].Value = 2.0;
            ws.Range["A3"].Value = 3.0;
            ws.Range["A4"].Value = 4.0;

            // Color A1 and A3 red (ColorIndex 3)
            ws.Range["A1"].Interior.ColorIndex = 3;
            ws.Range["A3"].Interior.ColorIndex = 3;

            EnterBacktickFormula(ws, "B1", "=`A1:A4.cells.where(c => c.color == 3).count()`");

            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(2.0, Convert.ToDouble(result));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Sum()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            ws.Range["A1"].Value = 100.0;
            ws.Range["A2"].Value = 200.0;
            ws.Range["A3"].Value = 300.0;

            EnterBacktickFormula(ws, "B1", "=`A1:A3.sum()`");

            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(600.0, Convert.ToDouble(result));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Where_GreaterThan()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            ws.Range["A1"].Value = 5.0;
            ws.Range["A2"].Value = 15.0;
            ws.Range["A3"].Value = 25.0;
            ws.Range["A4"].Value = 35.0;

            EnterBacktickFormula(ws, "B1", "=`A1:A4.where(x => x > 20).toArray()`");

            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B2 value: {ws.Range["B2"].Value}");

            Assert.NotNull(result);
            Assert.Equal(25.0, Convert.ToDouble(ws.Range["B1"].Value));
            Assert.Equal(35.0, Convert.ToDouble(ws.Range["B2"].Value));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_OrderBy()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            ws.Range["A1"].Value = 30.0;
            ws.Range["A2"].Value = 10.0;
            ws.Range["A3"].Value = 20.0;

            EnterBacktickFormula(ws, "B1", "=`A1:A3.orderBy(x => x).toArray()`");

            WaitForResult(ws, "B1");

            _output.WriteLine($"B1={ws.Range["B1"].Value}, B2={ws.Range["B2"].Value}, B3={ws.Range["B3"].Value}");

            Assert.Equal(10.0, Convert.ToDouble(ws.Range["B1"].Value));
            Assert.Equal(20.0, Convert.ToDouble(ws.Range["B2"].Value));
            Assert.Equal(30.0, Convert.ToDouble(ws.Range["B3"].Value));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void InvalidExpression_ShowsError()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            ws.Range["A1"].Value = 1.0;

            // Enter an invalid expression — should result in an error comment
            EnterBacktickFormula(ws, "B1", "=`A1.nonExistentMethod()`");

            // Wait a bit for the interceptor to process
            Thread.Sleep(5000);

            // The cell should either still contain the original text (interception failed to rewrite)
            // or have an error comment
            var formula = ws.Range["B1"].Formula2 as string;
            var value = ws.Range["B1"].Value as string;
            string? comment = null;
            try
            {
                comment = ws.Range["B1"].Comment?.Text() as string;
            }
            catch
            {
                // No comment
            }

            _output.WriteLine($"B1 formula: {formula}");
            _output.WriteLine($"B1 value: {value}");
            _output.WriteLine($"B1 comment: {comment}");

            // Either the cell wasn't rewritten (still has backticks) or there's an error comment
            var hasError = (formula?.Contains('`') ?? false) || (comment?.Contains("Error") ?? false);
            Assert.True(hasError, "Expected cell to show error for invalid expression");
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    /// <summary>
    ///     Enters a backtick formula into a cell as text.
    ///     Prefixes with apostrophe so Excel stores the =... literally as text.
    /// </summary>
    private static void EnterBacktickFormula(dynamic ws, string cellAddress, string formula)
    {
        var cell = ws.Range[cellAddress];
        cell.Value = "'" + formula;
    }

    /// <summary>
    ///     Polls a cell until the interceptor has rewritten it and the UDF has returned a result.
    /// </summary>
    private object? WaitForResult(dynamic ws, string cellAddress,
        int timeoutMs = 15000, int pollIntervalMs = 250)
    {
        var cell = ws.Range[cellAddress];
        var sw = Stopwatch.StartNew();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var formula = cell.Formula2 as string;
                object? value = cell.Value;

                if (formula != null && formula.StartsWith('=') && !formula.Contains('`'))
                {
                    if (value != null && value is not string)
                    {
                        return value;
                    }

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

        try
        {
            var finalFormula = cell.Formula2 as string;
            var finalValue = cell.Value;
            _output.WriteLine($"TIMEOUT for {cellAddress}. Formula2={finalFormula}, Value={finalValue}");
            return finalValue;
        }
        catch
        {
            return null;
        }
    }

    private static void CleanupWorksheet(dynamic ws)
    {
        try
        {
            ws.Delete();
        }
        catch
        {
            // Ignore
        }
    }
}
