using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class GroupByTests
{
    private static readonly Dictionary<string, int> ColumnMap =
        new(StringComparer.OrdinalIgnoreCase) { { "Category", 0 }, { "Amount", 1 } };

    private static RowCollection CreateTestCollection()
    {
        var rows = new Row[]
        {
            new(["A", 10.0], ColumnMap),
            new(["B", 20.0], ColumnMap),
            new(["A", 30.0], ColumnMap),
            new(["B", 40.0], ColumnMap),
            new(["A", 50.0], ColumnMap)
        };
        return new RowCollection(rows, ColumnMap);
    }

    [Fact]
    public void GroupBy_GroupsByKey()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        Assert.Equal(2, grouped.Count());
    }

    [Fact]
    public void GroupBy_RowGroupHasKey()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var first = grouped.First();
        Assert.Equal("A", first.Key);
    }

    [Fact]
    public void GroupBy_RowGroupHasCorrectRows()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var groupA = grouped.First(g => g.Key!.Equals("A"));
        Assert.Equal(3, groupA.Count());
    }

    [Fact]
    public void GroupBy_RowGroupInheritsWhere()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var groupA = grouped.First(g => g.Key!.Equals("A"));
        var filtered = groupA.Where(r => (double)r[1] > 20.0);
        Assert.Equal(2, filtered.Count());
    }

    [Fact]
    public void GroupBy_RowGroupInheritsSelect()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var groupA = grouped.First(g => g.Key!.Equals("A"));
        var selected = groupA.Select(r => r[1]);
        var result = selected.ToResult();
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
    }

    [Fact]
    public void GroupBy_Select_ProjectsEachGroup()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var result = grouped.Select(g => g.Count()).ToResult();
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0));
        Assert.Equal(3, arr[0, 0]); // A has 3 rows
        Assert.Equal(2, arr[1, 0]); // B has 2 rows
    }

    [Fact]
    public void GroupBy_Select_MultiColumn()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var result = grouped.Select(g => new object[] { g.Key, g.Count() }).ToResult();
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0));
        Assert.Equal(2, arr.GetLength(1));
        Assert.Equal("A", arr[0, 0]);
        Assert.Equal(3, arr[0, 1]);
        Assert.Equal("B", arr[1, 0]);
        Assert.Equal(2, arr[1, 1]);
    }

    [Fact]
    public void GroupBy_EmptyCollection()
    {
        var empty = new RowCollection(Array.Empty<Row>(), ColumnMap);
        var grouped = empty.GroupBy(r => r[0]);
        Assert.Equal(0, grouped.Count());
    }

    [Fact]
    public void GroupBy_PreservesColumnMap()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var group = grouped.First();
        // Column map is preserved — ToRange should work
        var range = group.ToRange();
        var result = range.ToResult();
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(1)); // 2 columns: Category, Amount
    }

    [Fact]
    public void GroupedRowCollection_Where()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var filtered = grouped.Where(g => g.Count() > 2);
        Assert.Equal(1, filtered.Count());
        Assert.Equal("A", filtered.First().Key);
    }

    [Fact]
    public void GroupedRowCollection_OrderBy()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var ordered = grouped.OrderByDescending(g => g.Count());
        Assert.Equal("A", ordered.First().Key); // A has 3, B has 2
    }

    [Fact]
    public void GroupBy_KeyUnwrapsColumnValue()
    {
        var grouped = CreateTestCollection().GroupBy(r => r["Category"]);
        var first = grouped.First();
        // Key should be unwrapped from ColumnValue to raw "A"
        Assert.IsType<string>(first.Key);
        Assert.Equal("A", first.Key);
    }

    [Fact]
    public void ResultConverter_GroupedRowCollection_FlattensWithKeyColumn()
    {
        var grouped = CreateTestCollection().GroupBy(r => r[0]);
        var result = ResultConverter.Convert(grouped);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(5, arr.GetLength(0)); // 3 + 2 rows
        Assert.Equal(3, arr.GetLength(1)); // key + 2 original columns
        Assert.Equal("A", arr[0, 0]);
        Assert.Equal(10.0, arr[0, 2]);
        Assert.Equal("B", arr[3, 0]);
    }

    [Fact]
    public void ResultConverter_EmptyGroupedRowCollection_ReturnsEmpty()
    {
        var empty = new RowCollection(Array.Empty<Row>(), ColumnMap);
        var grouped = empty.GroupBy(r => r[0]);
        var result = ResultConverter.Convert(grouped);
        Assert.Equal(string.Empty, result);
    }
}
