using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ResultConverterTests
{
    [Fact]
    public void ScalarToResult_Returns1x1Array()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.ToResult();
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal(1, result.GetLength(1));
        Assert.Equal(42.0, result[0, 0]);
    }

    [Fact]
    public void ArrayToResult_ReturnsSameArray()
    {
        var data = new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } };
        var arr = new ExcelArray(data);
        var result = arr.ToResult();
        Assert.Same(data, result);
    }

    [Fact]
    public void BoolToResult_Returns1x1Array()
    {
        var result = true.ToResult();
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal(1, result.GetLength(1));
        Assert.Equal(true, result[0, 0]);
    }

    [Fact]
    public void IntToResult_Returns1x1Array()
    {
        var result = 42.ToResult();
        Assert.Equal(42, result[0, 0]);
    }

    [Fact]
    public void DoubleToResult_Returns1x1Array()
    {
        var result = 3.14.ToResult();
        Assert.Equal(3.14, result[0, 0]);
    }

    [Fact]
    public void StringToResult_Returns1x1Array()
    {
        var result = "hello".ToResult();
        Assert.Equal("hello", result[0, 0]);
    }

    [Fact]
    public void IExcelRangeToResult_ConvertsFilteredRows()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        var filtered = arr.Where(r => (double)r[0] > 1.0);
        var result = filtered.ToResult();
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(2.0, result[0, 0]);
        Assert.Equal(3.0, result[1, 0]);
    }
}
