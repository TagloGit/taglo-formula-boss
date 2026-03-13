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
    public void MultipleBacktickExpressions_InOneFormula()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "B1", 3.0);
            TestUtilities.SetCellValue(ws, "B2", 7.0);

            // Two backtick expressions combined with +
            TestUtilities.EnterBacktickFormula(ws, "C1", "=`A1:A2.Sum()` + `B1:B2.Sum()`");

            var result = TestUtilities.WaitForResult(ws, "C1", _output);

            _output.WriteLine($"C1 formula: {TestUtilities.GetCellFormula(ws, "C1")}");
            _output.WriteLine($"C1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(40.0, Convert.ToDouble(result)); // 30 + 10
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void CellEscalation_RowBoldFilter()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Table with headers
            TestUtilities.SetCellValue(ws, "A1", "Name");
            TestUtilities.SetCellValue(ws, "B1", "Amount");
            TestUtilities.SetCellValue(ws, "A2", "Alice");
            TestUtilities.SetCellValue(ws, "B2", 100.0);
            TestUtilities.SetCellValue(ws, "A3", "Bob");
            TestUtilities.SetCellValue(ws, "B3", 200.0);
            TestUtilities.SetCellValue(ws, "A4", "Charlie");
            TestUtilities.SetCellValue(ws, "B4", 300.0);

            // Bold Bob's name
            TestUtilities.SetCellBold(ws, "A3", true);

            // Create table
            var range = ws.Range["A1:B4"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    // Filter rows where Name cell is bold
                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.Where(r => r[\"Name\"].Cell.Bold).Select(r => r[\"Amount\"]).Sum()`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1 value: {result}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(200.0, Convert.ToDouble(result)); // Only Bob's amount
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
    public void CellRgb_FilterEndToEnd()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);
            TestUtilities.SetCellValue(ws, "A4", 40.0);

            // Set A2 and A4 to red (RGB 255 = pure red in Excel's BGR format)
            TestUtilities.SetCellRgbColor(ws, "A2", 255);
            TestUtilities.SetCellRgbColor(ws, "A4", 255);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A4.Cells.Where(c => c.Rgb == 255).Sum()`");

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
    public void StatementBlock_CompilesAndExecutes()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            TestUtilities.EnterBacktickFormula(ws, "B1",
                "=`{ var total = A1:A3.Sum(); return total + 1; }`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B1 comment: {TestUtilities.GetCellComment(ws, "B1")}");

            Assert.NotNull(result);
            Assert.Equal(61.0, Convert.ToDouble(result)); // 10 + 20 + 30 + 1
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void RuntimeError_SurfacesAsValueError()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "not_a_number");

            // Sum on non-numeric data should cause a runtime error
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A1.Sum()`");

            // Wait for the interceptor to process
            Thread.Sleep(8000);

            var formula = TestUtilities.GetCellFormula(ws, "B1");
            var value = TestUtilities.GetCellValue(ws, "B1");

            _output.WriteLine($"B1 formula: {formula}");
            _output.WriteLine($"B1 value: {value} (type: {value?.GetType()?.Name})");

            // After rewrite, the UDF should return #VALUE! (represented as -2146826273 or COMException)
            // or the formula might not be rewritten at all
            Assert.NotNull(formula);

            // If rewritten, the value should be an error indicator
            if (formula!.StartsWith('=') && !formula.Contains('`'))
            {
                // UDF was registered — check for error value
                // Excel error values come through as int error codes
                var isError = value == null ||
                             value is int ||
                             (value is string s && s.StartsWith('#'));
                Assert.True(isError, $"Expected error value but got: {value}");
            }
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

    [Fact]
    public void LetFormula_ColumnAccessOnLetVariable_ShowsError()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up a table
            TestUtilities.SetCellValue(ws, "A1", "Name");
            TestUtilities.SetCellValue(ws, "B1", "Amount");
            TestUtilities.SetCellValue(ws, "A2", "Alice");
            TestUtilities.SetCellValue(ws, "B2", 100.0);
            TestUtilities.SetCellValue(ws, "A3", "Bob");
            TestUtilities.SetCellValue(ws, "B3", 200.0);

            var range = ws.Range["A1:B3"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    // LET formula where backtick expression uses column access on a LET variable
                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=LET(sales, {tableName}, `sales.Rows.Where(r => (double)r[\"Amount\"] > 150).Count()`)");

                    Thread.Sleep(5000);

                    var formula = TestUtilities.GetCellFormula(ws, "D1");
                    var comment = TestUtilities.GetCellComment(ws, "D1");

                    _output.WriteLine($"D1 formula: {formula}");
                    _output.WriteLine($"D1 comment: {comment}");

                    // Should show error about LET variable column access, not a COM exception
                    Assert.NotNull(comment);
                    Assert.Contains("LET variable", comment);
                    Assert.Contains("sales", comment);
                    // Formula should be left as backtick text (not rewritten)
                    Assert.True(formula?.Contains('`') ?? false, "Formula should not have been rewritten");
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
    public void RowPath_Aggregate_Sum()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "Item");
            TestUtilities.SetCellValue(ws, "A2", "A");
            TestUtilities.SetCellValue(ws, "A3", "B");
            TestUtilities.SetCellValue(ws, "A4", "C");
            TestUtilities.SetCellValue(ws, "B1", "Price");
            TestUtilities.SetCellValue(ws, "B2", 10.0);
            TestUtilities.SetCellValue(ws, "B3", 20.0);
            TestUtilities.SetCellValue(ws, "B4", 30.0);

            var range = ws.Range["A1:B4"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.Aggregate(0.0, (acc, r) => acc + (double)r[\"Price\"])`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1 value: {result}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(60.0, Convert.ToDouble(result));
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
    public void RowPath_Scan_RunningTotal()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "Item");
            TestUtilities.SetCellValue(ws, "A2", "A");
            TestUtilities.SetCellValue(ws, "A3", "B");
            TestUtilities.SetCellValue(ws, "A4", "C");
            TestUtilities.SetCellValue(ws, "B1", "Price");
            TestUtilities.SetCellValue(ws, "B2", 10.0);
            TestUtilities.SetCellValue(ws, "B3", 20.0);
            TestUtilities.SetCellValue(ws, "B4", 30.0);

            var range = ws.Range["A1:B4"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.Scan(0.0, (acc, r) => acc + (double)r[\"Price\"])`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1={TestUtilities.GetCellValue(ws, "D1")}, D2={TestUtilities.GetCellValue(ws, "D2")}, D3={TestUtilities.GetCellValue(ws, "D3")}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(10.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "D1")));
                    Assert.Equal(30.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "D2")));
                    Assert.Equal(60.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "D3")));
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
    public void RowPath_GroupBy_SelectCount()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "Category");
            TestUtilities.SetCellValue(ws, "A2", "Fruit");
            TestUtilities.SetCellValue(ws, "A3", "Veg");
            TestUtilities.SetCellValue(ws, "A4", "Fruit");
            TestUtilities.SetCellValue(ws, "A5", "Veg");
            TestUtilities.SetCellValue(ws, "A6", "Fruit");
            TestUtilities.SetCellValue(ws, "B1", "Price");
            TestUtilities.SetCellValue(ws, "B2", 1.0);
            TestUtilities.SetCellValue(ws, "B3", 2.0);
            TestUtilities.SetCellValue(ws, "B4", 3.0);
            TestUtilities.SetCellValue(ws, "B5", 4.0);
            TestUtilities.SetCellValue(ws, "B6", 5.0);

            var range = ws.Range["A1:B6"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.GroupBy(r => r[\"Category\"]).Select(g => g.Count())`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1={TestUtilities.GetCellValue(ws, "D1")}, D2={TestUtilities.GetCellValue(ws, "D2")}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(3.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "D1"))); // Fruit: 3
                    Assert.Equal(2.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "D2"))); // Veg: 2
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
    public void ReEdit_LET_ChangedExpression_RecompilesUdf()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            // Step 1: Enter initial LET formula
            TestUtilities.EnterBacktickFormula(ws, "B1",
                "=LET(x, `A1:A3.Sum()`, x)");

            var result1 = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"Step 1 - B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"Step 1 - B1 value: {result1}");

            Assert.NotNull(result1);
            Assert.Equal(60.0, Convert.ToDouble(result1)); // 10 + 20 + 30

            // Step 2: Re-enter with a different expression (simulating editor re-edit)
            TestUtilities.EnterBacktickFormula(ws, "B1",
                "=LET(x, `A1:A3.Count()`, x)");

            var result2 = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"Step 2 - B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"Step 2 - B1 value: {result2}");

            Assert.NotNull(result2);
            Assert.Equal(3.0, Convert.ToDouble(result2)); // Count = 3
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ReEdit_SimpleBacktick_ChangedExpression_RecompilesUdf()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            // Step 1: Enter initial backtick formula
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Sum()`");

            var result1 = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"Step 1 - B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"Step 1 - B1 value: {result1}");

            Assert.NotNull(result1);
            Assert.Equal(60.0, Convert.ToDouble(result1));

            // Step 2: Re-enter with a different expression
            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.Count()`");

            var result2 = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"Step 2 - B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"Step 2 - B1 value: {result2}");

            Assert.NotNull(result2);
            Assert.Equal(3.0, Convert.ToDouble(result2));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void RowPath_GroupBy_SelectKeyAndCount()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "Category");
            TestUtilities.SetCellValue(ws, "A2", "Fruit");
            TestUtilities.SetCellValue(ws, "A3", "Veg");
            TestUtilities.SetCellValue(ws, "A4", "Fruit");
            TestUtilities.SetCellValue(ws, "B1", "Price");
            TestUtilities.SetCellValue(ws, "B2", 10.0);
            TestUtilities.SetCellValue(ws, "B3", 20.0);
            TestUtilities.SetCellValue(ws, "B4", 30.0);

            var range = ws.Range["A1:B4"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.GroupBy(r => r[\"Category\"]).Select(g => new object[] {{ g.Key, g.Count() }})`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1={TestUtilities.GetCellValue(ws, "D1")}, E1={TestUtilities.GetCellValue(ws, "E1")}");
                    _output.WriteLine($"D2={TestUtilities.GetCellValue(ws, "D2")}, E2={TestUtilities.GetCellValue(ws, "E2")}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal("Fruit", TestUtilities.GetCellValue(ws, "D1"));
                    Assert.Equal(2.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "E1")));
                    Assert.Equal("Veg", TestUtilities.GetCellValue(ws, "D2"));
                    Assert.Equal(1.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "E2")));
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
    public void Foreach_SumsRange()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            TestUtilities.EnterBacktickFormula(ws, "B1",
                "=LET(data, A1:A3, result, `{ var sum = 0.0; foreach (var x in data) { sum += (double)x; } return sum; }`, result)");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1 value: {result}");
            _output.WriteLine($"B1 comment: {TestUtilities.GetCellComment(ws, "B1")}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void CellsWhereBold_ReturnsValues()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);
            TestUtilities.SetCellValue(ws, "A4", 40.0);

            // Bold A2 and A4
            TestUtilities.SetCellBold(ws, "A2", true);
            TestUtilities.SetCellBold(ws, "A4", true);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A4.Cells.Where(c => c.Bold)`");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);

            _output.WriteLine($"B1 formula: {TestUtilities.GetCellFormula(ws, "B1")}");
            _output.WriteLine($"B1={TestUtilities.GetCellValue(ws, "B1")}, B2={TestUtilities.GetCellValue(ws, "B2")}");
            _output.WriteLine($"B1 comment: {TestUtilities.GetCellComment(ws, "B1")}");

            Assert.NotNull(result);
            Assert.Equal(20.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B1")));
            Assert.Equal(40.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "B2")));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void SingleRow_ReturnsHorizontalArray()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", "Name");
            TestUtilities.SetCellValue(ws, "B1", "Amount");
            TestUtilities.SetCellValue(ws, "A2", "Alice");
            TestUtilities.SetCellValue(ws, "B2", 100.0);
            TestUtilities.SetCellValue(ws, "A3", "Bob");
            TestUtilities.SetCellValue(ws, "B3", 200.0);

            var range = ws.Range["A1:B3"];
            var tables = ws.ListObjects;
            try
            {
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}.Rows.First()`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1={TestUtilities.GetCellValue(ws, "D1")}, E1={TestUtilities.GetCellValue(ws, "E1")}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal("Alice", TestUtilities.GetCellValue(ws, "D1"));
                    Assert.Equal(100.0, Convert.ToDouble(TestUtilities.GetCellValue(ws, "E1")));
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
    public void ValuePath_ToList_UnwrapsExcelValues()
    {
        // Regression: List<ExcelValue> hit the generic IEnumerable handler which
        // didn't unwrap ExcelValue → Excel got wrapper objects → #VALUE!
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);

            TestUtilities.EnterBacktickFormula(ws, "B1", "=`A1:A3.ToList()`");

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
    public void CrossSheetRangeRef_SumsCorrectly()
    {
        // Create a second worksheet with data, then reference it from the first
        var dataWs = _excel.AddWorksheet();
        var formulaWs = _excel.AddWorksheet();
        try
        {
            // Get the data sheet name for the cross-sheet reference
            string sheetName = dataWs.Name;

            // Put data on the data sheet
            TestUtilities.SetCellValue(dataWs, "A1", 10.0);
            TestUtilities.SetCellValue(dataWs, "A2", 20.0);
            TestUtilities.SetCellValue(dataWs, "A3", 30.0);

            // Enter a backtick formula on the formula sheet referencing the data sheet
            TestUtilities.EnterBacktickFormula(formulaWs, "A1",
                $"=`{sheetName}!A1:A3.Sum()`");

            var result = TestUtilities.WaitForResult(formulaWs, "A1", _output);

            _output.WriteLine($"A1 formula: {TestUtilities.GetCellFormula(formulaWs, "A1")}");
            _output.WriteLine($"A1 value: {result}");

            Assert.NotNull(result);
            Assert.Equal(60.0, Convert.ToDouble(result));
        }
        finally
        {
            TestUtilities.CleanupWorksheet(formulaWs);
            TestUtilities.CleanupWorksheet(dataWs);
        }
    }

    [Fact]
    public void Table_ColumnAccess_Sum()
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
                var table = tables.Add(1, range, Type.Missing, 1);
                try
                {
                    var tableName = (string)table.Name;

                    // Use table column bracket access with Sum
                    TestUtilities.EnterBacktickFormula(ws, "D1",
                        $"=`{tableName}[\"Amount\"].Sum()`");

                    var result = TestUtilities.WaitForResult(ws, "D1", _output);

                    _output.WriteLine($"D1 formula: {TestUtilities.GetCellFormula(ws, "D1")}");
                    _output.WriteLine($"D1 value: {result}");
                    _output.WriteLine($"D1 comment: {TestUtilities.GetCellComment(ws, "D1")}");

                    Assert.NotNull(result);
                    Assert.Equal(300.0, Convert.ToDouble(result)); // 100 + 200
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
}
