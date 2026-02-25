using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class RowTests
{
    [Fact]
    public void IndexAccess_ReturnsColumnValue()
    {
        var row = new Row(["a", 2, true], null);
        Assert.Equal("a", row[0]);
        Assert.Equal(2.0, (double)row[1]);
        Assert.True((bool)row[2]);
    }

    [Fact]
    public void NegativeIndex_CountsFromEnd()
    {
        var row = new Row(["first", "middle", "last"], null);
        Assert.Equal("last", row[-1]);
        Assert.Equal("middle", row[-2]);
    }

    [Fact]
    public void NamedAccess_UsesColumnMap()
    {
        var map = new Dictionary<string, int> { ["Name"] = 0, ["Age"] = 1 };
        var row = new Row(["Alice", 30.0], map);

        Assert.Equal("Alice", row["Name"]);
        Assert.Equal(30.0, (double)row["Age"]);
    }

    [Fact]
    public void NamedAccess_ThrowsForUnknownColumn()
    {
        var map = new Dictionary<string, int> { ["Name"] = 0 };
        var row = new Row(["Alice"], map);

        Assert.Throws<KeyNotFoundException>(() => row["Missing"]);
    }

    [Fact]
    public void DynamicAccess_ResolvesColumnNames()
    {
        var map = new Dictionary<string, int> { ["Name"] = 0, ["Age"] = 1 };
        dynamic row = new Row(["Alice", 30.0], map);

        ColumnValue name = row.Name;
        ColumnValue age = row.Age;

        Assert.Equal("Alice", name);
        Assert.Equal(30.0, (double)age);
    }

    [Fact]
    public void ColumnCount_ReturnsCorrectCount()
    {
        var row = new Row([1, 2, 3], null);
        Assert.Equal(3, row.ColumnCount);
    }
}
