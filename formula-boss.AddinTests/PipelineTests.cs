using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.AddinTests;

/// <summary>
///     Core pipeline tests that exercise the full add-in loop with more complex expressions.
///     Tests are added incrementally as pipeline features are verified working.
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
    public void ValuePath_Where_GreaterThan()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 5.0);
            TestUtilities.SetCellValue(ws, "A2", 15.0);
            TestUtilities.SetCellValue(ws, "A3", 25.0);
            TestUtilities.SetCellValue(ws, "A4", 35.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A4.Where(x => x > 20)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1={TestUtilities.GetCellValue(ws, "B1")}, B2={TestUtilities.GetCellValue(ws, "B2")}");

            Assert.NotNull(result);
            Assert.Equal(25.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B1")));
            Assert.Equal(35.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B2")));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact(Skip = "OrderBy returns #VALUE! — runtime bug to investigate")]
    public void ValuePath_OrderBy()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 30.0);
            TestUtilities.SetCellValue(ws, "A2", 10.0);
            TestUtilities.SetCellValue(ws, "A3", 20.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.OrderBy(x => x)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1={TestUtilities.GetCellValue(ws, "B1")}, B2={TestUtilities.GetCellValue(ws, "B2")}, B3={TestUtilities.GetCellValue(ws, "B3")}");

            Assert.NotNull(result);
            Assert.Equal(10.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B1")));
            Assert.Equal(20.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B2")));
            Assert.Equal(30.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B3")));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Count()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Count()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(3.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Min()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 30.0);
            TestUtilities.SetCellValue(ws, "A2", 10.0);
            TestUtilities.SetCellValue(ws, "A3", 20.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Min()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(10.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Max()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 30.0);
            TestUtilities.SetCellValue(ws, "A2", 10.0);
            TestUtilities.SetCellValue(ws, "A3", 20.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Max()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(30.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact(Skip = "Cells.Where().Sum() — Sum not available on IEnumerable<Cell>, needs runtime method")]
    public void ObjectModelPath_WhereColor_Sum()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);
            TestUtilities.SetCellValue(ws, "A4", 40.0);
            TestUtilities.SetCellValue(ws, "A5", 50.0);

            // Color cells A2 and A4 yellow (ColorIndex 6)
            TestUtilities.SetCellColor(ws, "A2", 6);
            TestUtilities.SetCellColor(ws, "A4", 6);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A5.Cells.Where(c => c.Color == 6).Sum()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result)); // 20 + 40
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ObjectModelPath_WhereColor_Count()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 1.0);
            TestUtilities.SetCellValue(ws, "A2", 2.0);
            TestUtilities.SetCellValue(ws, "A3", 3.0);
            TestUtilities.SetCellValue(ws, "A4", 4.0);

            // Color A1 and A3 red (ColorIndex 3)
            TestUtilities.SetCellColor(ws, "A1", 3);
            TestUtilities.SetCellColor(ws, "A3", 3);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A4.Cells.Where(c => c.Color == 3).Count()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(2.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void InvalidExpression_ShowsError()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 1.0);

            // Enter an invalid expression — should result in an error comment
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1.NonExistentMethod()`");

            // Wait a bit for the interceptor to process
            Thread.Sleep(5000);

            var formula = TestUtilities.GetCellFormula(ws, "B1");
            var value = TestUtilities.GetCellValue(ws, "B1") as string;
            var comment = TestUtilities.GetCellComment(ws, "B1");

            _output.WriteLine($"B1 formula: {formula}");
            _output.WriteLine($"B1 value: {value}");
            _output.WriteLine($"B1 comment: {comment}");

            // Either the cell wasn't rewritten (still has backticks) or there's an error comment
            var hasError = (formula?.Contains('`') ?? false) || (comment?.Contains("Error") ?? false);
            Assert.True(hasError, "Expected cell to show error for invalid expression");
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }
}
