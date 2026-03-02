using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ResultConverterTests
{
    [Fact]
    public void ScalarToResult_ReturnsBareValue()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.ToResult();
        Assert.Equal(42.0, result);
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
    public void BoolToResult_ReturnsBareValue()
    {
        var result = true.ToResult();
        Assert.Equal(true, result);
    }

    [Fact]
    public void IntToResult_ReturnsBareValue()
    {
        var result = 42.ToResult();
        Assert.Equal(42, result);
    }

    [Fact]
    public void DoubleToResult_ReturnsBareValue()
    {
        var result = 3.14.ToResult();
        Assert.Equal(3.14, result);
    }

    [Fact]
    public void StringToResult_ReturnsBareValue()
    {
        var result = "hello".ToResult();
        Assert.Equal("hello", result);
    }

    [Fact]
    public void IExcelRangeToResult_ConvertsFilteredRows()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        var filtered = arr.Where(r => (double)r[0] > 1.0);
        var result = filtered.ToResult();
        var resultArr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, resultArr.GetLength(0));
        Assert.Equal(2.0, resultArr[0, 0]);
        Assert.Equal(3.0, resultArr[1, 0]);
    }

    [Fact]
    public void Convert_NullReturnsEmpty()
    {
        Assert.Equal(string.Empty, ResultConverter.Convert(null));
    }

    [Fact]
    public void Convert_ScalarReturnsBareValue()
    {
        Assert.Equal(42.0, ResultConverter.Convert(42.0));
        Assert.Equal("hello", ResultConverter.Convert("hello"));
        Assert.Equal(true, ResultConverter.Convert(true));
    }

    [Fact]
    public void Convert_ExcelValueDelegates()
    {
        var scalar = new ExcelScalar(99.0);
        Assert.Equal(99.0, ResultConverter.Convert(scalar));
    }
}
