using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ColumnCollectionTests
{
    private static ExcelArray MakeArray() => new(
        new object?[,]
        {
            { 1.0, "A", true },
            { 2.0, "B", false },
            { 3.0, "C", true }
        });

    private static ExcelTable MakeTable() => new(
        new object?[,]
        {
            { "Alice", 30.0, "NY" },
            { "Bob", 25.0, "LA" },
            { "Carol", 35.0, "NY" }
        },
        new[] { "Name", "Age", "City" });

    // --- ExcelArray.Cols ---

    [Fact]
    public void Array_Cols_ReturnsCorrectCount()
    {
        var arr = MakeArray();
        var cols = arr.Cols;
        Assert.Equal(3, cols.Count());
    }

    [Fact]
    public void Array_Cols_HasNullNames()
    {
        var arr = MakeArray();
        foreach (var col in arr.Cols)
        {
            Assert.Null(col.Name);
        }
    }

    [Fact]
    public void Array_Cols_HasCorrectIndex()
    {
        var arr = MakeArray();
        var cols = arr.Cols.ToList();
        Assert.Equal(0, cols[0].Index);
        Assert.Equal(1, cols[1].Index);
        Assert.Equal(2, cols[2].Index);
    }

    [Fact]
    public void Array_Cols_ContainsCorrectValues()
    {
        var arr = MakeArray();
        var cols = arr.Cols.ToList();

        // First column: 1, 2, 3
        Assert.Equal(3, cols[0].RowCount);
        Assert.Equal(1.0, cols[0][0, 0].RawValue);
        Assert.Equal(2.0, cols[0][1, 0].RawValue);
        Assert.Equal(3.0, cols[0][2, 0].RawValue);

        // Second column: A, B, C
        Assert.Equal("A", cols[1][0, 0].RawValue);
        Assert.Equal("B", cols[1][1, 0].RawValue);
        Assert.Equal("C", cols[1][2, 0].RawValue);
    }

    // --- ExcelTable.Cols ---

    [Fact]
    public void Table_Cols_ReturnsCorrectCount()
    {
        var table = MakeTable();
        Assert.Equal(3, table.Cols.Count());
    }

    [Fact]
    public void Table_Cols_HasCorrectNames()
    {
        var table = MakeTable();
        var cols = table.Cols.ToList();
        Assert.Equal("Name", cols[0].Name);
        Assert.Equal("Age", cols[1].Name);
        Assert.Equal("City", cols[2].Name);
    }

    [Fact]
    public void Table_Cols_HasCorrectIndex()
    {
        var table = MakeTable();
        var cols = table.Cols.ToList();
        Assert.Equal(0, cols[0].Index);
        Assert.Equal(1, cols[1].Index);
        Assert.Equal(2, cols[2].Index);
    }

    [Fact]
    public void Table_Cols_ColumnSum()
    {
        var table = MakeTable();
        var cols = table.Cols.ToList();
        Assert.Equal(90.0, (double)cols[1].Sum());
    }

    // --- ExcelScalar.Cols ---

    [Fact]
    public void Scalar_Cols_ReturnsSingleColumn()
    {
        var scalar = new ExcelScalar(42.0);
        var cols = scalar.Cols;
        Assert.Equal(1, cols.Count());
    }

    [Fact]
    public void Scalar_Cols_HasNullName()
    {
        var scalar = new ExcelScalar(42.0);
        var col = scalar.Cols.First(c => true);
        Assert.Null(col.Name);
    }

    [Fact]
    public void Scalar_Cols_HasCorrectValue()
    {
        var scalar = new ExcelScalar(42.0);
        var col = scalar.Cols.First(c => true);
        Assert.Equal(42.0, col[0, 0].RawValue);
    }

    // --- ColumnCollection methods ---

    [Fact]
    public void Where_FiltersColumns()
    {
        var table = MakeTable();
        var filtered = table.Cols.Where(c => c.Name == "Age");
        Assert.Equal(1, filtered.Count());
    }

    [Fact]
    public void Select_ProjectsColumns()
    {
        var table = MakeTable();
        var result = table.Cols.Select(c => c.Name);
        Assert.Equal(3, result.Count());
    }

    [Fact]
    public void First_ReturnsMatchingColumn()
    {
        var table = MakeTable();
        var col = table.Cols.First(c => c.Name == "City");
        Assert.Equal("City", col.Name);
        Assert.Equal(2, col.Index);
    }

    [Fact]
    public void FirstOrDefault_ReturnsNullWhenNotFound()
    {
        var table = MakeTable();
        var col = table.Cols.FirstOrDefault(c => c.Name == "NonExistent");
        Assert.Null(col);
    }

    [Fact]
    public void OrderBy_SortsColumns()
    {
        var table = MakeTable();
        var sorted = table.Cols.OrderBy(c => c.Name).ToList();
        Assert.Equal("Age", sorted[0].Name);
        Assert.Equal("City", sorted[1].Name);
        Assert.Equal("Name", sorted[2].Name);
    }

    [Fact]
    public void Take_ReturnsFirstN()
    {
        var table = MakeTable();
        var first2 = table.Cols.Take(2).ToList();
        Assert.Equal(2, first2.Count);
        Assert.Equal("Name", first2[0].Name);
        Assert.Equal("Age", first2[1].Name);
    }

    [Fact]
    public void Skip_SkipsFirstN()
    {
        var table = MakeTable();
        var last2 = table.Cols.Skip(1).ToList();
        Assert.Equal(2, last2.Count);
        Assert.Equal("Age", last2[0].Name);
        Assert.Equal("City", last2[1].Name);
    }

    // --- Row.Index ---

    [Fact]
    public void Row_Index_SetByArray()
    {
        var arr = MakeArray();
        var rows = arr.Rows.ToList();
        Assert.Equal(0, rows[0].Index);
        Assert.Equal(1, rows[1].Index);
        Assert.Equal(2, rows[2].Index);
    }

    [Fact]
    public void Row_Index_DefaultsToZero()
    {
        var row = new Row(new object?[] { 1.0, 2.0 }, null);
        Assert.Equal(0, row.Index);
    }

    // --- Column.Index (renamed from ColumnIndex) ---

    [Fact]
    public void Column_Index_RenamedFromColumnIndex()
    {
        var col = new Column(new object?[,] { { 1.0 }, { 2.0 } }, "Test", 5);
        Assert.Equal(5, col.Index);
    }

    // --- Column.Name nullable ---

    [Fact]
    public void Column_Name_CanBeNull()
    {
        var col = new Column(new object?[,] { { 1.0 } }, null, 0);
        Assert.Null(col.Name);
    }

    [Fact]
    public void Column_Name_CanBeSet()
    {
        var col = new Column(new object?[,] { { 1.0 } }, "MyCol", 0);
        Assert.Equal("MyCol", col.Name);
    }
}
