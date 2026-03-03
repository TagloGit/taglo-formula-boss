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

    [Fact]
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

    [Fact]
    public void ValuePath_Select_Multiply()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 3.0);
            TestUtilities.SetCellValue(ws, "A2", 5.0);
            TestUtilities.SetCellValue(ws, "A3", 7.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Select(x => x * 2)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 comment: {TestUtilities.GetCellComment(ws, "B1")}");
            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1={TestUtilities.GetCellValue(ws, "B1")}, B2={TestUtilities.GetCellValue(ws, "B2")}, B3={TestUtilities.GetCellValue(ws, "B3")}");

            Assert.NotNull(result);
            Assert.Equal(6.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B1")));
            Assert.Equal(10.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B2")));
            Assert.Equal(14.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B3")));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Any_True()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 5.0);
            TestUtilities.SetCellValue(ws, "A2", 15.0);
            TestUtilities.SetCellValue(ws, "A3", 25.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Any(x => x > 20)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.True(Convert.ToBoolean(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_Any_False()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 5.0);
            TestUtilities.SetCellValue(ws, "A2", 10.0);
            TestUtilities.SetCellValue(ws, "A3", 15.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Any(x => x > 20)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.False(Convert.ToBoolean(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_All_True()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 25.0);
            TestUtilities.SetCellValue(ws, "A2", 30.0);
            TestUtilities.SetCellValue(ws, "A3", 35.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.All(x => x > 20)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.True(Convert.ToBoolean(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_All_False()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 5.0);
            TestUtilities.SetCellValue(ws, "A2", 30.0);
            TestUtilities.SetCellValue(ws, "A3", 35.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.All(x => x > 20)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.False(Convert.ToBoolean(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ValuePath_First()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 5.0);
            TestUtilities.SetCellValue(ws, "A2", 15.0);
            TestUtilities.SetCellValue(ws, "A3", 25.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.First(x => x > 10)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(15.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void LetFormula_TwoRange()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Range 1: A1:A3
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            // Range 2: B1:B3
            TestUtilities.SetCellValue(ws, "B1", 1.0);
            TestUtilities.SetCellValue(ws, "B2", 2.0);
            TestUtilities.SetCellValue(ws, "B3", 3.0);

            // LET formula with two ranges: sum of each
            TestUtilities.EnterBacktickFormula(ws, "C1",
                "=LET(a,A1:A3,b,B1:B3,`a.Sum()` + `b.Sum()`)");

            var result = TestUtilities.WaitForResult(ws, "C1", _output);

            _output.WriteLine($"C1 formula: {TestUtilities.GetCellFormula(ws, "C1")}");
            _output.WriteLine($"C1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(66.0, Convert.ToDouble(result)); // 60 + 6
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void LetFormula_VariableReusedInLambda()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Threshold in A1
            TestUtilities.SetCellValue(ws, "A1", 15.0);

            // Data in B1:B4
            TestUtilities.SetCellValue(ws, "B1", 5.0);
            TestUtilities.SetCellValue(ws, "B2", 20.0);
            TestUtilities.SetCellValue(ws, "B3", 10.0);
            TestUtilities.SetCellValue(ws, "B4", 25.0);

            // LET with threshold variable used in lambda
            TestUtilities.EnterBacktickFormula(ws, "C1",
                "=LET(threshold,A1,data,B1:B4,`data.Where(x => x > threshold).Count()`)");

            var result = TestUtilities.WaitForResult(ws, "C1", _output);

            _output.WriteLine($"C1 formula: {TestUtilities.GetCellFormula(ws, "C1")}");
            _output.WriteLine($"C1 value: {result}");
            _output.WriteLine($"C1 comment: {TestUtilities.GetCellComment(ws, "C1")}");

            Assert.NotNull(result);
            Assert.Equal(2.0, Convert.ToDouble(result)); // 20 and 25
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
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
    public void Table_BracketColumnAccess()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up a table with headers
            TestUtilities.SetCellValue(ws, "A1", "Name");
            TestUtilities.SetCellValue(ws, "A2", "Alice");
            TestUtilities.SetCellValue(ws, "A3", "Bob");
            TestUtilities.SetCellValue(ws, "B1", "Amount");
            TestUtilities.SetCellValue(ws, "B2", 100.0);
            TestUtilities.SetCellValue(ws, "B3", 200.0);

            // Create an Excel ListObject (table) from the range
            var range = ws.Range["A1:B3"];
            var tables = ws.ListObjects;
            try
            {
                // xlSrcRange = 1, xlYes = 1
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    // Use table reference with bracket column access
                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.Where(r => (double)r[\"Amount\"] > 150).Count()`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1 value: {result}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(1.0, Convert.ToDouble(result)); // Only Bob (200)
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            }
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void EmptyRange_ReturnsZeroCount()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Empty cells
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Count()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B1 comment: {TestUtilities.GetCellComment(ws, "B1")}");

            Assert.NotNull(result);
            Assert.Equal(3.0, Convert.ToDouble(result)); // 3 cells even if empty
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ScalarResult_ReturnsBareValue()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 42.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A1.Sum()`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(42.0, Convert.ToDouble(result));
            // Should NOT spill to B2
            Assert.Null(TestUtilities.GetCellValue(ws, "B2"));
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
