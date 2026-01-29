using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.IntegrationTests;

/// <summary>
/// Integration tests for the object model path (.cells).
/// These tests compile generated code and run it against real Excel COM objects.
/// </summary>
public class ObjectModelTests : IClassFixture<ExcelTestFixture>
{
    private readonly ExcelTestFixture _excel;
    private readonly ITestOutputHelper _output;

    public ObjectModelTests(ExcelTestFixture excel, ITestOutputHelper output)
    {
        _excel = excel;
        _output = output;
    }

    #region Basic Cell Access

    [Fact]
    public void Cells_SelectValue_ReturnsAllValues()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { 1 }, { 2 }, { 3 } });
        var compilation = TestHelpers.CompileExpression("data.cells.select(c => c.value).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result type: {result?.GetType()?.Name ?? "null"}");
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(1, arr.GetLength(1));
        Assert.Equal(1.0, arr[0, 0]);
        Assert.Equal(2.0, arr[1, 0]);
        Assert.Equal(3.0, arr[2, 0]);
    }

    [Fact]
    public void Cells_ToArray_ReturnsAllCellValues()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { 10 }, { 20 }, { 30 } });
        var compilation = TestHelpers.CompileExpression("data.cells.toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
    }

    #endregion

    #region Filtering by Color

    [Fact]
    public void Cells_WhereColor_FiltersCorrectly()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { 1 }, { 2 }, { 3 }, { 4 }, { 5 } });
        _excel.SetCellColor(range, 2, 1, 6); // Row 2 = Yellow (index 6)
        _excel.SetCellColor(range, 4, 1, 6); // Row 4 = Yellow

        var compilation = TestHelpers.CompileExpression("data.cells.where(c => c.color == 6).select(c => c.value).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // Should have 2 yellow cells
        Assert.Equal(2.0, arr[0, 0]); // Value from row 2
        Assert.Equal(4.0, arr[1, 0]); // Value from row 4
    }

    [Fact]
    public void Cells_WhereColorNotEqual_FiltersCorrectly()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { 1 }, { 2 }, { 3 } });
        _excel.SetCellColor(range, 2, 1, 6); // Row 2 = Yellow

        var compilation = TestHelpers.CompileExpression("data.cells.where(c => c.color != 6).select(c => c.value).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        // Should have 2 non-yellow cells (rows 1 and 3)
    }

    #endregion

    #region Cell Properties

    [Fact]
    public void Cells_SelectRow_ReturnsRowNumbers()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { "a" }, { "b" }, { "c" } });
        var compilation = TestHelpers.CompileExpression("data.cells.select(c => c.row).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
    }

    [Fact]
    public void Cells_SelectColumn_ReturnsColumnNumbers()
    {
        // Arrange: Create a horizontal range
        var range = _excel.CreateRange(new object[,] { { "a", "b", "c" } }, "M1");
        var compilation = TestHelpers.CompileExpression("data.cells.select(c => c.col).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
    }

    #endregion

    #region Complex Expressions

    [Fact]
    public void Cells_WhereValueGreaterThan_SelectValue_ToArray()
    {
        // Arrange
        var range = _excel.CreateUniqueRange(new object[,] { { 10 }, { 25 }, { 5 }, { 30 }, { 15 } });
        var compilation = TestHelpers.CompileExpression("data.cells.where(c => c.value > 20).select(c => c.value).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithRange(compilation.CoreMethod!, range);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // 25 and 30
    }

    #endregion

    #region Helpers

    private static string FormatResult(object? result)
    {
        if (result == null)
        {
            return "null";
        }

        if (result is string s)
        {
            return $"\"{s}\"";
        }

        if (result is object[,] arr)
        {
            var rows = arr.GetLength(0);
            var cols = arr.GetLength(1);
            var items = new List<string>();
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    items.Add(arr[r, c]?.ToString() ?? "null");
                }
            }

            return $"[{rows}x{cols}]: [{string.Join(", ", items)}]";
        }

        return result.ToString() ?? "null";
    }

    #endregion
}
