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
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Sum()`");

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
}
