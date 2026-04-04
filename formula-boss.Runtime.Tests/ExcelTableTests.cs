using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelTableTests
{
    private static ExcelTable MakeTable() => new(
        new object?[,]
        {
            { "Alice", 30.0, "Engineering" }, { "Bob", 25.0, "Marketing" }, { "Charlie", 35.0, "Engineering" }
        },
        ["Name", "Age", "Department"]);

    [Fact]
    public void Headers_ReturnsColumnNames()
    {
        var table = MakeTable();
        Assert.Equal(["Name", "Age", "Department"], table.Headers);
    }

    [Fact]
    public void Rows_HaveNamedColumnAccess()
    {
        var table = MakeTable();
        var rows = table.Rows.ToList();
        Assert.Equal("Alice", rows[0]["Name"]);
        Assert.Equal(30.0, (double)rows[0]["Age"]);
    }

    [Fact]
    public void Rows_Where_FiltersWithNamedAccess()
    {
        var table = MakeTable();
        var result = table.Rows.Where(r => (string)r["Department"] == "Engineering");
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Rows_Where_FilterWithDynamicAccess()
    {
        var table = MakeTable();
        var result = table.Rows.Where(r => (double)r["Age"] > 28.0);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void ColumnAccess_CaseInsensitive()
    {
        var table = MakeTable();
        var rows = table.Rows.ToList();
        Assert.Equal("Alice", rows[0]["name"]);
        Assert.Equal("Alice", rows[0]["NAME"]);
    }

    [Fact]
    public void Rows_Select_MapsRowsToValues()
    {
        var table = MakeTable();
        var result = table.Rows.Select(r => new ExcelScalar((string)r["Name"]));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("Alice", rows[0][0].RawValue);
    }

    [Fact]
    public void Rows_OrderBy_SortsByColumn()
    {
        var table = MakeTable();
        var sorted = table.Rows.OrderBy(r => (double)r["Age"]);
        var rows = sorted.ToList();
        Assert.Equal("Bob", rows[0]["Name"].RawValue);
        Assert.Equal("Alice", rows[1]["Name"].RawValue);
        Assert.Equal("Charlie", rows[2]["Name"].RawValue);
    }

    // --- Column indexer tests ---

    [Fact]
    public void StringIndexer_ReturnsColumn()
    {
        var table = MakeTable();
        var col = table["Name"];
        Assert.IsType<Column>(col);
        Assert.Equal("Name", col.Name);
        Assert.Equal(0, col.Index);
    }

    [Fact]
    public void StringIndexer_ReturnsCorrectValues()
    {
        var table = MakeTable();
        var col = table["Age"];
        Assert.Equal(3, col.RowCount);
        Assert.Equal(1, col.ColCount);
        Assert.Equal(30.0, col[0, 0].RawValue);
        Assert.Equal(25.0, col[1, 0].RawValue);
        Assert.Equal(35.0, col[2, 0].RawValue);
    }

    [Fact]
    public void StringIndexer_CaseInsensitive()
    {
        var table = MakeTable();
        var col = table["age"];
        Assert.Equal(3, col.RowCount);
        Assert.Equal(30.0, col[0, 0].RawValue);
    }

    [Fact]
    public void StringIndexer_InvalidColumn_Throws()
    {
        var table = MakeTable();
        Assert.Throws<KeyNotFoundException>(() => table["NonExistent"]);
    }

    [Fact]
    public void StringIndexer_ColumnSum()
    {
        var table = MakeTable();
        var sum = table["Age"].Sum();
        Assert.Equal(90.0, (double)sum);
    }

    [Fact]
    public void StringIndexer_ColumnCount()
    {
        var table = MakeTable();
        Assert.Equal(3, table["Age"].Count());
    }

    [Fact]
    public void StringIndexer_EmptyTable()
    {
        var table = new ExcelTable(new object?[0, 2], ["A", "B"]);
        var col = table["A"];
        Assert.Equal(0, col.RowCount);
        Assert.Equal(0, col.Count());
    }

    // --- Lookup tests ---

    [Fact]
    public void Lookup_ExactMatch_ReturnsValue()
    {
        var table = MakeTable();
        var result = ExcelTable.Lookup("Alice", table["Name"], table["Age"]);
        Assert.Equal(30.0, result.RawValue);
    }

    [Fact]
    public void Lookup_CaseInsensitiveStringMatch()
    {
        var table = MakeTable();
        var result = ExcelTable.Lookup("alice", table["Name"], table["Age"]);
        Assert.Equal(30.0, result.RawValue);
    }

    [Fact]
    public void Lookup_NumericMatch()
    {
        var table = MakeTable();
        var result = ExcelTable.Lookup(30.0, table["Age"], table["Name"]);
        Assert.Equal("Alice", result.RawValue);
    }

    [Fact]
    public void Lookup_NoMatch_Throws()
    {
        var table = MakeTable();
        Assert.Throws<KeyNotFoundException>(() => ExcelTable.Lookup("Nobody", table["Name"], table["Age"]));
    }

    [Fact]
    public void Lookup_NoMatch_ReturnsIfNotFound()
    {
        var table = MakeTable();
        var result = ExcelTable.Lookup("Nobody", table["Name"], table["Age"], "N/A");
        Assert.Equal("N/A", result.RawValue);
    }

    [Fact]
    public void Lookup_ReturnsFirstOccurrence()
    {
        var table = new ExcelTable(
            new object?[,] { { "Engineering", 1.0 }, { "Marketing", 2.0 }, { "Engineering", 3.0 } },
            ["Department", "Id"]);
        var result = ExcelTable.Lookup("Engineering", table["Department"], table["Id"]);
        Assert.Equal(1.0, result.RawValue);
    }

    [Fact]
    public void Lookup_ExcelValueUnwrapped()
    {
        var table = MakeTable();
        var result = ExcelTable.Lookup(new ExcelScalar("Bob"), table["Name"], table["Age"]);
        Assert.Equal(25.0, result.RawValue);
    }

    [Fact]
    public void Lookup_MismatchedColumnLengths_Throws()
    {
        var table = MakeTable();
        var shortTable = new ExcelTable(
            new object?[,] { { "X", 1.0 } }, ["Name", "Age"]);
        Assert.Throws<ArgumentException>(() => ExcelTable.Lookup("Alice", table["Name"], shortTable["Age"]));
    }
}
