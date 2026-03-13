using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ColumnTests
{
    private static ExcelTable MakeTable() => new(
        new object?[,]
        {
            { "Alice", 30.0, "Engineering" },
            { "Bob", 25.0, "Marketing" },
            { "Charlie", 35.0, "Engineering" }
        },
        ["Name", "Age", "Department"]);

    [Fact]
    public void StringIndexer_ReturnsColumn()
    {
        var table = MakeTable();
        var col = table["Name"];

        Assert.IsType<Column>(col);
        Assert.Equal("Name", col.Name);
        Assert.Equal(0, col.ColumnIndex);
    }

    [Fact]
    public void Column_HasCorrectRowCount()
    {
        var table = MakeTable();
        var col = table["Age"];

        Assert.Equal(3, col.Count());
    }

    [Fact]
    public void Column_ValuesMatchSourceColumn()
    {
        var table = MakeTable();
        var col = table["Name"];
        var values = col.Select(r => r[0]).ToList();

        Assert.Equal(3, values.Count);
    }

    [Fact]
    public void Column_RowsAreSingleCell()
    {
        var table = MakeTable();
        var col = table["Age"];
        var rows = col.ToList();

        Assert.Equal(1, rows[0].ColumnCount);
        Assert.Equal(30.0, rows[0][0].Value);
        Assert.Equal(25.0, rows[1][0].Value);
        Assert.Equal(35.0, rows[2][0].Value);
    }

    [Fact]
    public void Column_CaseInsensitive()
    {
        var table = MakeTable();
        var col = table["name"];

        Assert.Equal("name", col.Name);
        Assert.Equal(0, col.ColumnIndex);
    }

    [Fact]
    public void Column_ThrowsOnInvalidName()
    {
        var table = MakeTable();
        Assert.Throws<KeyNotFoundException>(() => table["NonExistent"]);
    }

    [Fact]
    public void Column_Where_Filters()
    {
        var table = MakeTable();
        var col = table["Age"];
        var filtered = col.Where(r => (double)r[0] > 28.0);

        Assert.Equal(2, filtered.Count());
    }

    [Fact]
    public void Column_OrderBy_Sorts()
    {
        var table = MakeTable();
        var col = table["Age"];
        var sorted = col.OrderBy(r => (double)r[0]);
        var rows = sorted.ToList();

        Assert.Equal(25.0, rows[0][0].Value);
        Assert.Equal(30.0, rows[1][0].Value);
        Assert.Equal(35.0, rows[2][0].Value);
    }

    [Fact]
    public void Column_NamedAccess_WorksViaSingleColMap()
    {
        var table = MakeTable();
        var col = table["Name"];
        var rows = col.ToList();

        Assert.Equal("Alice", rows[0]["Name"].Value);
    }

    [Fact]
    public void Lookup_FindsMatch()
    {
        var table = MakeTable();
        var result = table.Lookup("Bob", table["Name"], table["Age"]);

        Assert.Equal(25.0, result);
    }

    [Fact]
    public void Lookup_ReturnsIfNotFound_WhenNoMatch()
    {
        var table = MakeTable();
        var result = table.Lookup("Nobody", table["Name"], table["Age"], -1.0);

        Assert.Equal(-1.0, result);
    }

    [Fact]
    public void Lookup_ReturnsNull_WhenNoMatchAndNoDefault()
    {
        var table = MakeTable();
        var result = table.Lookup("Nobody", table["Name"], table["Age"]);

        Assert.Null(result);
    }

    [Fact]
    public void Lookup_FirstMatchWins()
    {
        var table = new ExcelTable(
            new object?[,]
            {
                { "Engineering", 1.0 },
                { "Marketing", 2.0 },
                { "Engineering", 3.0 }
            },
            ["Department", "Id"]);

        var result = table.Lookup("Engineering", table["Department"], table["Id"]);

        Assert.Equal(1.0, result);
    }

    [Fact]
    public void Lookup_StringResult()
    {
        var table = MakeTable();
        var result = table.Lookup(30.0, table["Age"], table["Name"]);

        Assert.Equal("Alice", result);
    }
}
