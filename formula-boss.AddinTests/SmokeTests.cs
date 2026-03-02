using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.AddinTests;

/// <summary>
///     Smoke tests that verify the Formula Boss XLL loads in Excel and basic interception works.
///     Uses custom COM automation to launch Excel with the add-in registered.
/// </summary>
[Collection("Excel Addin")]
public class SmokeTests
{
    private readonly ExcelAddinFixture _excel;
    private readonly ITestOutputHelper _output;

    public SmokeTests(ExcelAddinFixture excel, ITestOutputHelper output)
    {
        _excel = excel;
        _output = output;
    }

    [Fact]
    public void AddinLoadsSuccessfully()
    {
        // If we get here, Excel launched and the XLL was registered via RegisterXLL.
        Assert.NotNull(_excel.Application);
        _output.WriteLine("Excel launched and add-in registered successfully.");
    }

    [Fact]
    public void ScalarExpressionReturnsCorrectValue()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up source data
            ws.Range["A1"].Value = 10.0;
            ws.Range["A2"].Value = 20.0;
            ws.Range["A3"].Value = 30.0;

            // Enter a backtick formula.
            // The interceptor fires on SheetChange for cells containing =...`...`
            // We need to enter it as text (apostrophe prefix), which Excel stores as =`...`
            EnterBacktickFormula(ws, "B1", "=`A1:A3.sum()`");

            // Wait for interception + Roslyn compilation + UDF execution
            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 formula: {ws.Range["B1"].Formula2}");
            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ArrayExpressionSpillsCorrectly()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up source data
            ws.Range["A1"].Value = 1.0;
            ws.Range["A2"].Value = 2.0;
            ws.Range["A3"].Value = 3.0;

            // Enter a backtick formula that returns an array
            EnterBacktickFormula(ws, "B1", "=`A1:A3.toArray()`");

            // Wait for interception to rewrite and UDF to execute
            var result = WaitForResult(ws, "B1");

            _output.WriteLine($"B1 formula: {ws.Range["B1"].Formula2}");
            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B2 value: {ws.Range["B2"].Value}");
            _output.WriteLine($"B3 value: {ws.Range["B3"].Value}");

            Assert.NotNull(result);

            // Check spilled values
            Assert.Equal(1.0, Convert.ToDouble(ws.Range["B1"].Value));
            Assert.Equal(2.0, Convert.ToDouble(ws.Range["B2"].Value));
            Assert.Equal(3.0, Convert.ToDouble(ws.Range["B3"].Value));
        }
        finally
        {
            CleanupWorksheet(ws);
        }
    }

    /// <summary>
    ///     Enters a backtick formula into a cell as text.
    ///     Prefixes with apostrophe so Excel stores the =... literally as text,
    ///     which is how users actually enter backtick formulas.
    /// </summary>
    private static void EnterBacktickFormula(dynamic ws, string cellAddress, string formula)
    {
        var cell = ws.Range[cellAddress];
        // Prefix with apostrophe — Excel hides it but stores the value as text
        // The interceptor sees the cell text starting with = and containing backticks
        cell.Value = "'" + formula;
    }

    /// <summary>
    ///     Polls a cell until the interceptor has rewritten it from text to a UDF call
    ///     and the UDF has returned a result, or until timeout.
    /// </summary>
    private object? WaitForResult(dynamic ws, string cellAddress,
        int timeoutMs = 15000, int pollIntervalMs = 250)
    {
        var cell = ws.Range[cellAddress];
        var sw = System.Diagnostics.Stopwatch.StartNew();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                string? formula = cell.Formula2 as string;
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
            _output.WriteLine($"TIMEOUT waiting for {cellAddress}. Formula2={finalFormula}, Value={finalValue}");
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
            // Ignore cleanup errors
        }
    }
}
