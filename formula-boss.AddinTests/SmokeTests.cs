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
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            // Enter a backtick formula
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.sum()`");

            // Wait for interception + Roslyn compilation + UDF execution
            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ArrayExpressionSpillsCorrectly()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up source data
            TestUtilities.SetCellValue(ws, "A1", 1.0);
            TestUtilities.SetCellValue(ws, "A2", 2.0);
            TestUtilities.SetCellValue(ws, "A3", 3.0);

            // Enter a backtick formula that returns an array
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.toArray()`");

            // Wait for interception to rewrite and UDF to execute
            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B2 value: {TestUtilities.GetCellValue(ws, "B2")}");
            _output.WriteLine($"B3 value: {TestUtilities.GetCellValue(ws, "B3")}");

            Assert.NotNull(result);

            // Check spilled values
            Assert.Equal(1.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B1")));
            Assert.Equal(2.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B2")));
            Assert.Equal(3.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B3")));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }
}
