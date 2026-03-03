using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelValueTests
{
    [Fact]
    public void Wrap_Scalar_ReturnsExcelScalar()
    {
        var result = ExcelValue.Wrap(42.0);
        Assert.IsType<ExcelScalar>(result);
        Assert.Equal(42.0, result.RawValue);
    }

    [Fact]
    public void Wrap_String_ReturnsExcelScalar()
    {
        var result = ExcelValue.Wrap("hello");
        Assert.IsType<ExcelScalar>(result);
    }

    [Fact]
    public void Wrap_Bool_ReturnsExcelScalar()
    {
        var result = ExcelValue.Wrap(true);
        Assert.IsType<ExcelScalar>(result);
    }

    [Fact]
    public void Wrap_Null_ReturnsExcelScalar()
    {
        var result = ExcelValue.Wrap(null);
        Assert.IsType<ExcelScalar>(result);
        Assert.Null(result.RawValue);
    }

    [Fact]
    public void Wrap_Array_ReturnsExcelArray()
    {
        var result = ExcelValue.Wrap(new object?[,] { { 1.0, 2.0 } });
        Assert.IsType<ExcelArray>(result);
    }

    [Fact]
    public void Wrap_MultiRowArray_ReturnsExcelArray()
    {
        var result = ExcelValue.Wrap(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        Assert.IsType<ExcelArray>(result);
    }

    [Fact]
    public void Wrap_SingleCellArray_ReturnsExcelScalar()
    {
        // Single-cell ExcelReference values are returned as 1x1 arrays by GetValuesFromReference.
        // Wrap should unwrap these to ExcelScalar so they work correctly in lambda comparisons.
        var result = ExcelValue.Wrap(new object[,] { { 15.0 } });
        Assert.IsType<ExcelScalar>(result);
        Assert.Equal(15.0, result.RawValue);
    }

    [Fact]
    public void Wrap_SingleCellArrayWithHeaders_ReturnsExcelTable()
    {
        // 1x1 with headers should still be a table, not unwrapped
        var result = ExcelValue.Wrap(new object[,] { { "Val" } }, new[] { "Col" });
        Assert.IsType<ExcelTable>(result);
    }

    [Fact]
    public void Wrap_ArrayWithHeaders_ReturnsExcelTable()
    {
        var result = ExcelValue.Wrap(
            new object?[,] { { 1.0, 2.0 } },
            new[] { "A", "B" });
        Assert.IsType<ExcelTable>(result);
    }

    [Fact]
    public void Wrap_ExcelValue_ReturnsSameInstance()
    {
        var scalar = new ExcelScalar(42.0);
        var result = ExcelValue.Wrap(scalar);
        Assert.Same(scalar, result);
    }

    [Fact]
    public void Equality_WorksBetweenExcelValues()
    {
        var a = new ExcelScalar(42.0);
        var b = new ExcelScalar(42.0);
        var c = new ExcelScalar(99.0);

        Assert.True(a == b);
        Assert.False(a == c);
        Assert.True(a != c);
    }

    [Fact]
    public void ComparisonOperators_WorkBetweenExcelValues()
    {
        var a = new ExcelScalar(10.0);
        var b = new ExcelScalar(20.0);

        Assert.True(a < b);
        Assert.True(b > a);
        Assert.True(a <= b);
        Assert.True(b >= a);
    }

    [Fact]
    public void ComparisonOperators_WorkWithDoubles()
    {
        var a = new ExcelScalar(10.0);

        Assert.True(a > 5.0);
        Assert.True(a < 15.0);
        Assert.True(5.0 < a);
        Assert.True(15.0 > a);
    }

    [Fact]
    public void LetScenario_MultiCellRange_WhereWithScalar_Counts()
    {
        // Simulates the LET scenario: data=B1:B4 (multi-cell), threshold=A1 (single-cell)
        // Both arrive via GetValuesFromReference as object[,]
        var data = ExcelValue.Wrap(new object[,] { { 5.0 }, { 20.0 }, { 10.0 }, { 25.0 } });
        var threshold = ExcelValue.Wrap(new object[,] { { 15.0 } }); // 1x1 from single-cell ref

        Assert.IsType<ExcelArray>(data);
        Assert.IsType<ExcelScalar>(threshold); // must unwrap to scalar

        var count = data.Where(x => x > threshold).Count();
        Assert.Equal(2, count); // 20 and 25
    }

    [Fact]
    public void LetScenario_MultiCellRange_WhereWithScalar_Sum()
    {
        var data = ExcelValue.Wrap(new object[,] { { 5.0 }, { 20.0 }, { 10.0 }, { 25.0 } });
        var threshold = ExcelValue.Wrap(new object[,] { { 15.0 } });

        var result = data.Where(x => x > threshold).Sum();
        Assert.Equal(45.0, (double)result); // 20 + 25
    }
}
