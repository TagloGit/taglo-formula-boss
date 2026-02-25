using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelTableTests
{
    private static ExcelTable MakeTable() => new(
        new object?[,]
        {
            { "Alice", 30.0, "Engineering" },
            { "Bob", 25.0, "Marketing" },
            { "Charlie", 35.0, "Engineering" }
        },
        new[] { "Name", "Age", "Department" });

    [Fact]
    public void Headers_ReturnsColumnNames()
    {
        var table = MakeTable();
        Assert.Equal(new[] { "Name", "Age", "Department" }, table.Headers);
    }

    [Fact]
    public void Rows_HaveNamedColumnAccess()
    {
        var table = MakeTable();
        var rows = table.Rows.ToList();
        Assert.Equal("Alice", (string?)rows[0]["Name"]);
        Assert.Equal(30.0, (double)rows[0]["Age"]);
    }

    [Fact]
    public void Rows_HaveDynamicColumnAccess()
    {
        var table = MakeTable();
        dynamic row = table.Rows.First();
        Assert.Equal("Alice", (string?)(ColumnValue)row.Name);
    }

    [Fact]
    public void Where_FiltersWithNamedAccess()
    {
        var table = MakeTable();
        var result = table.Where(r => (string?)r["Department"] == "Engineering");
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Where_FilterWithDynamicAccess()
    {
        var table = MakeTable();
        var result = table.Where(r => (double)r["Age"] > 28.0);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void ColumnAccess_CaseInsensitive()
    {
        var table = MakeTable();
        var row = table.Rows.First();
        Assert.Equal("Alice", (string?)row["name"]);
        Assert.Equal("Alice", (string?)row["NAME"]);
    }

    [Fact]
    public void Select_MapsRowsToValues()
    {
        var table = MakeTable();
        var result = table.Select(r => new ExcelScalar((string?)r["Name"]));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("Alice", rows[0][0].Value);
    }

    [Fact]
    public void OrderBy_SortsByColumn()
    {
        var table = MakeTable();
        var result = table.OrderBy(r => (double)r["Age"]);
        var rows = result.Rows.ToList();
        Assert.Equal("Bob", rows[0]["Name"].Value);
        Assert.Equal("Alice", rows[1]["Name"].Value);
        Assert.Equal("Charlie", rows[2]["Name"].Value);
    }
}
