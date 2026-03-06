using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class RowCollectionTests
{
    private static readonly Dictionary<string, int> ColumnMap =
        new(StringComparer.OrdinalIgnoreCase) { { "Name", 0 }, { "Price", 1 } };

    private static RowCollection CreateTestCollection()
    {
        var rows = new Row[]
        {
            new(["A", 10.0], ColumnMap),
            new(["B", 20.0], ColumnMap),
            new(["C", 30.0], ColumnMap)
        };
        return new RowCollection(rows, ColumnMap);
    }

    [Fact]
    public void Aggregate_FoldsOverRows()
    {
        var result = CreateTestCollection().Aggregate(0.0, (acc, r) => acc + (double)r[1]);
        Assert.Equal(60.0, (double)result);
    }

    [Fact]
    public void Aggregate_WithStringSeed()
    {
        var result = CreateTestCollection().Aggregate("", (acc, r) => acc + (string)r[0]);
        Assert.Equal("ABC", (string)result);
    }

    [Fact]
    public void Aggregate_EmptyCollection_ReturnsSeed()
    {
        var empty = new RowCollection(Array.Empty<Row>(), ColumnMap);
        var result = empty.Aggregate(42.0, (acc, r) => acc + (double)r[1]);
        Assert.Equal(42.0, (double)result);
    }

    [Fact]
    public void Scan_RunningFold()
    {
        var result = CreateTestCollection().Scan(0.0, (acc, r) => acc + (double)r[1]);
        var arr = (object?[,])result.ToResult();
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(1, arr.GetLength(1));
        Assert.Equal(10.0, arr[0, 0]);
        Assert.Equal(30.0, arr[1, 0]);
        Assert.Equal(60.0, arr[2, 0]);
    }

    [Fact]
    public void Scan_EmptyCollection_ReturnsEmpty()
    {
        var empty = new RowCollection(Array.Empty<Row>(), ColumnMap);
        var result = empty.Scan(0.0, (acc, r) => acc + (double)r[1]);
        var arr = (object?[,])result.ToResult();
        Assert.Equal(0, arr.GetLength(0));
    }
}
