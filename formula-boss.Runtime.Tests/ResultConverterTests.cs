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
        var filtered = arr.Where(v => (double)v > 1.0);
        var result = filtered.ToResult();
        var resultArr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, resultArr.GetLength(0));
        Assert.Equal(2.0, resultArr[0, 0]);
        Assert.Equal(3.0, resultArr[1, 0]);
    }

    [Fact]
    public void Convert_NullReturnsEmpty() => Assert.Equal(string.Empty, ResultConverter.Convert(null));

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

    [Fact]
    public void Convert_EnumerableRow_ReturnsObjectArray()
    {
        var columnMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { { "A", 0 }, { "B", 1 } };
        IEnumerable<Row> rows = new[] { new Row([1.0, "X"], columnMap), new Row([2.0, "Y"], columnMap) };

        var result = ResultConverter.Convert(rows);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0));
        Assert.Equal(2, arr.GetLength(1));
        Assert.Equal(1.0, arr[0, 0]);
        Assert.Equal("Y", arr[1, 1]);
    }

    [Fact]
    public void Convert_EnumerableColumnValue_ReturnsVerticalArray()
    {
        IEnumerable<ColumnValue> values = new[] { new ColumnValue(10.0), new ColumnValue(20.0), new ColumnValue(30.0) };

        var result = ResultConverter.Convert(values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(1, arr.GetLength(1));
        Assert.Equal(10.0, arr[0, 0]);
        Assert.Equal(30.0, arr[2, 0]);
    }

    [Fact]
    public void Convert_RowResult_SpillsAsSingleRow()
    {
        var row = new Row(["Alice", 30.0], null);
        ExcelValue rowAsExcel = row; // implicit conversion to ExcelArray 1×2
        var result = ResultConverter.Convert(rowAsExcel);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(1, arr.GetLength(0));
        Assert.Equal(2, arr.GetLength(1));
        Assert.Equal("Alice", arr[0, 0]);
        Assert.Equal(30.0, arr[0, 1]);
    }

    [Fact]
    public void Convert_ColumnValue_ReturnsSingleValue()
    {
        // A single ColumnValue isn't IEnumerable<ColumnValue>, so it falls through to generic return
        // But wrapping in an enumerable of one tests the path
        var result = ResultConverter.Convert(new[] { new ColumnValue(42.0) });
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(1, arr.GetLength(0));
        Assert.Equal(42.0, arr[0, 0]);
    }
}
