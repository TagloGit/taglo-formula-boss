using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ColumnTests
{
    private static Column MakeColumn() => new(
        new object?[,] { { 10.0 }, { 20.0 }, { 30.0 } },
        "Score", 2);

    [Fact]
    public void Name_ReturnsColumnName()
    {
        var col = MakeColumn();
        Assert.Equal("Score", col.Name);
    }

    [Fact]
    public void ColumnIndex_ReturnsIndex()
    {
        var col = MakeColumn();
        Assert.Equal(2, col.ColumnIndex);
    }

    [Fact]
    public void RowCount_MatchesData()
    {
        var col = MakeColumn();
        Assert.Equal(3, col.RowCount);
    }

    [Fact]
    public void ColCount_IsOne()
    {
        var col = MakeColumn();
        Assert.Equal(1, col.ColCount);
    }

    [Fact]
    public void Sum_ReturnsTotal()
    {
        var col = MakeColumn();
        Assert.Equal(60.0, (double)col.Sum());
    }

    [Fact]
    public void Count_ReturnsElementCount()
    {
        var col = MakeColumn();
        Assert.Equal(3, col.Count());
    }

    [Fact]
    public void Average_ReturnsAverage()
    {
        var col = MakeColumn();
        Assert.Equal(20.0, (double)col.Average());
    }

    [Fact]
    public void Min_ReturnsMinimum()
    {
        var col = MakeColumn();
        Assert.Equal(10.0, (double)col.Min());
    }

    [Fact]
    public void Max_ReturnsMaximum()
    {
        var col = MakeColumn();
        Assert.Equal(30.0, (double)col.Max());
    }

    [Fact]
    public void Where_FiltersValues()
    {
        var col = MakeColumn();
        var result = col.Where(x => (double)x > 15.0);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Select_TransformsValues()
    {
        var col = MakeColumn();
        var result = col.Select(x => new ExcelScalar((double)x * 2));
        Assert.Equal(3, result.Count());
    }

    [Fact]
    public void Indexer_AccessesElements()
    {
        var col = MakeColumn();
        Assert.Equal(10.0, col[0, 0].RawValue);
        Assert.Equal(20.0, col[1, 0].RawValue);
        Assert.Equal(30.0, col[2, 0].RawValue);
    }

    [Fact]
    public void Rows_ReturnsRowCollection()
    {
        var col = MakeColumn();
        var rows = col.Rows.ToList();
        Assert.Equal(3, rows.Count);
    }
}
