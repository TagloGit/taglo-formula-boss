using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelArrayTests
{
    private static ExcelArray MakeArray() => new(new object?[,]
    {
        { 1.0, "Alice" },
        { 2.0, "Bob" },
        { 3.0, "Charlie" }
    });

    [Fact]
    public void Rows_IteratesAllRows()
    {
        var arr = MakeArray();
        var rows = arr.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal("Charlie", (string?)rows[2][1]);
    }

    [Fact]
    public void Count_ReturnsRowCount()
    {
        Assert.Equal(3, MakeArray().Count());
    }

    [Fact]
    public void Where_FiltersRows()
    {
        var arr = MakeArray();
        var result = arr.Where(r => (double)r[0] > 1.0);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Any_TrueWhenMatching()
    {
        var arr = MakeArray();
        Assert.True(arr.Any(r => (string?)r[1] == "Bob"));
        Assert.False(arr.Any(r => (string?)r[1] == "Dave"));
    }

    [Fact]
    public void All_TrueWhenAllMatch()
    {
        var arr = MakeArray();
        Assert.True(arr.All(r => (double)r[0] > 0));
        Assert.False(arr.All(r => (double)r[0] > 2));
    }

    [Fact]
    public void First_ReturnsFirstMatch_SingleColumn()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        var result = arr.First(r => (double)r[0] > 1.0);
        Assert.Equal(2.0, (double)result);
    }

    [Fact]
    public void First_ReturnsFirstMatch_MultiColumn()
    {
        var arr = MakeArray();
        var result = arr.First(r => (double)r[0] > 1.0);
        var resultArray = (ExcelArray)result;
        var row = resultArray.Rows.First();
        Assert.Equal(2.0, (double)row[0]);
        Assert.Equal("Bob", (string?)row[1]);
    }

    [Fact]
    public void First_ThrowsWhenNoMatch()
    {
        Assert.Throws<InvalidOperationException>(() =>
            MakeArray().First(r => (double)r[0] > 100));
    }

    [Fact]
    public void FirstOrDefault_ReturnsNullWhenNoMatch()
    {
        var result = MakeArray().FirstOrDefault(r => (double)r[0] > 100);
        Assert.Null(result);
    }

    [Fact]
    public void Select_FlattensTo1D()
    {
        var arr = MakeArray();
        var result = arr.Select(r => new ExcelScalar((double)r[0] * 10));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(10.0, (double)rows[0][0]);
        Assert.Equal(30.0, (double)rows[2][0]);
    }

    [Fact]
    public void OrderBy_SortsRows()
    {
        var arr = new ExcelArray(new object?[,]
        {
            { 3.0 }, { 1.0 }, { 2.0 }
        });
        var result = arr.OrderBy(r => (double)r[0]);
        var rows = result.Rows.ToList();
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
        Assert.Equal(3.0, (double)rows[2][0]);
    }

    [Fact]
    public void OrderByDescending_SortsRowsDescending()
    {
        var arr = new ExcelArray(new object?[,]
        {
            { 1.0 }, { 3.0 }, { 2.0 }
        });
        var result = arr.OrderByDescending(r => (double)r[0]);
        var rows = result.Rows.ToList();
        Assert.Equal(3.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
        Assert.Equal(1.0, (double)rows[2][0]);
    }

    [Fact]
    public void Take_PositiveCount_TakesFromStart()
    {
        var result = MakeArray().Take(2);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Take_NegativeCount_TakesFromEnd()
    {
        var result = MakeArray().Take(-1);
        var rows = result.Rows.ToList();
        Assert.Single(rows);
        Assert.Equal(3.0, (double)rows[0][0]);
    }

    [Fact]
    public void Skip_PositiveCount_SkipsFromStart()
    {
        var result = MakeArray().Skip(2);
        Assert.Equal(1, result.Count());
    }

    [Fact]
    public void Skip_NegativeCount_SkipsFromEnd()
    {
        var result = MakeArray().Skip(-1);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Distinct_RemovesDuplicateRows()
    {
        var arr = new ExcelArray(new object?[,]
        {
            { 1.0, "a" },
            { 2.0, "b" },
            { 1.0, "a" }
        });
        var result = arr.Distinct();
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Sum_SumsAllValues()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        Assert.Equal(6.0, (double)arr.Sum());
    }

    [Fact]
    public void Min_ReturnsMinValue()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        Assert.Equal(1.0, (double)arr.Min());
    }

    [Fact]
    public void Max_ReturnsMaxValue()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        Assert.Equal(3.0, (double)arr.Max());
    }

    [Fact]
    public void Average_ReturnsAverageOfAllValues()
    {
        var arr = new ExcelArray(new object?[,] { { 2.0 }, { 4.0 }, { 6.0 } });
        Assert.Equal(4.0, (double)arr.Average());
    }

    [Fact]
    public void Aggregate_FoldsOverRows()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        var result = arr.Aggregate(
            new ExcelScalar(0.0),
            (acc, row) => new ExcelScalar((double)acc + (double)row[0]));
        Assert.Equal(6.0, (double)result);
    }

    [Fact]
    public void Scan_RunningFold()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        var result = arr.Scan(
            new ExcelScalar(0.0),
            (acc, row) => new ExcelScalar((double)acc + (double)row[0]));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[1][0]);
        Assert.Equal(6.0, (double)rows[2][0]);
    }

    [Fact]
    public void SelectMany_FlattensResults()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 } });
        var result = arr.SelectMany(r =>
            new ExcelValue[] { new ExcelScalar((double)r[0]), new ExcelScalar((double)r[0] * 10) });
        Assert.Equal(4, result.Count());
    }

    [Fact]
    public void EmptyArray_CountReturnsZero()
    {
        var arr = new ExcelArray(new object?[0, 2]);
        Assert.Equal(0, arr.Count());
    }

    [Fact]
    public void EmptyArray_AnyReturnsFalse()
    {
        var arr = new ExcelArray(new object?[0, 1]);
        Assert.False(arr.Any(_ => true));
    }

    [Fact]
    public void EmptyArray_AllReturnsTrue()
    {
        var arr = new ExcelArray(new object?[0, 1]);
        Assert.True(arr.All(_ => false));
    }
}
