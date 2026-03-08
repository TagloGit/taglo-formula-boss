using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelArrayTests
{
    private static ExcelArray MakeArray() =>
        new(new object?[,] { { 1.0, "Alice" }, { 2.0, "Bob" }, { 3.0, "Charlie" } });

    private static ExcelArray MakeSingleColumn() =>
        new(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });

    // --- Rows (RowCollection) ---

    [Fact]
    public void Rows_IteratesAllRows()
    {
        var arr = MakeArray();
        var rows = arr.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal("Charlie", rows[2][1]);
    }

    [Fact]
    public void Rows_ReturnsRowCollection()
    {
        var arr = MakeArray();
        var rc = arr.Rows;
        Assert.IsType<RowCollection>(rc);
        Assert.Equal(3, rc.Count());
    }

    // --- Element-wise Count ---

    [Fact]
    public void Count_ReturnsTotalCellCount() => Assert.Equal(6, MakeArray().Count());

    [Fact]
    public void Count_SingleColumn() => Assert.Equal(3, MakeSingleColumn().Count());

    // --- Element-wise Where ---

    [Fact]
    public void Where_FiltersCellsElementWise()
    {
        var arr = MakeSingleColumn();
        var result = arr.Where(v => (double)v > 1.0);
        Assert.Equal(2, result.Count());
    }

    // --- Element-wise Any/All ---

    [Fact]
    public void Any_TrueWhenMatchingCell()
    {
        var arr = MakeArray();
        Assert.True(arr.Any(v => (string?)v == "Bob"));
        Assert.False(arr.Any(v => (string?)v == "Dave"));
    }

    [Fact]
    public void All_TrueWhenAllCellsMatch()
    {
        var arr = MakeSingleColumn();
        Assert.True(arr.All(v => (double)v > 0));
        Assert.False(arr.All(v => (double)v > 2));
    }

    // --- Element-wise First ---

    [Fact]
    public void First_ReturnsFirstMatchingCell()
    {
        var arr = MakeSingleColumn();
        var result = arr.First(v => (double)v > 1.0);
        Assert.Equal(2.0, (double)result);
    }

    [Fact]
    public void First_ThrowsWhenNoMatch()
    {
        Assert.Throws<InvalidOperationException>(() =>
            MakeArray().First(v => (string?)v == "Nobody"));
    }

    [Fact]
    public void FirstOrDefault_ReturnsNullWhenNoMatch()
    {
        var result = MakeSingleColumn().FirstOrDefault(v => (double)v > 100);
        Assert.Null(result);
    }

    // --- Element-wise Select ---

    [Fact]
    public void Select_TransformsCellsElementWise()
    {
        var arr = MakeSingleColumn();
        var result = arr.Select(v => new ExcelScalar((double)v * 10));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(10.0, (double)rows[0][0]);
        Assert.Equal(30.0, (double)rows[2][0]);
    }

    // --- Element-wise OrderBy ---

    [Fact]
    public void OrderBy_IdentitySelector_SortsByComparable()
    {
        var arr = new ExcelArray(new object?[,] { { 30.0 }, { 10.0 }, { 20.0 } });
        var result = arr.OrderBy(x => x);
        var rows = result.Rows.ToList();
        Assert.Equal(10.0, (double)rows[0][0]);
        Assert.Equal(20.0, (double)rows[1][0]);
        Assert.Equal(30.0, (double)rows[2][0]);
    }

    [Fact]
    public void OrderBy_SortsCellsElementWise()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        var result = arr.OrderBy(v => (double)v);
        var rows = result.Rows.ToList();
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
        Assert.Equal(3.0, (double)rows[2][0]);
    }

    [Fact]
    public void OrderByDescending_SortsCellsDescending()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 3.0 }, { 2.0 } });
        var result = arr.OrderByDescending(v => (double)v);
        var rows = result.Rows.ToList();
        Assert.Equal(3.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
        Assert.Equal(1.0, (double)rows[2][0]);
    }

    // --- Partitioning ---

    [Fact]
    public void Take_PositiveCount_TakesFromStart()
    {
        var result = MakeSingleColumn().Take(2);
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Take_NegativeCount_TakesFromEnd()
    {
        var result = MakeSingleColumn().Take(-1);
        var rows = result.Rows.ToList();
        Assert.Single(rows);
        Assert.Equal(3.0, (double)rows[0][0]);
    }

    [Fact]
    public void Skip_PositiveCount_SkipsFromStart()
    {
        var result = MakeSingleColumn().Skip(2);
        Assert.Equal(1, result.Count());
    }

    [Fact]
    public void Skip_NegativeCount_SkipsFromEnd()
    {
        var result = MakeSingleColumn().Skip(-1);
        Assert.Equal(2, result.Count());
    }

    // --- Distinct ---

    [Fact]
    public void Distinct_RemovesDuplicateCells()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 }, { 1.0 } });
        var result = arr.Distinct();
        Assert.Equal(2, result.Count());
    }

    // --- Aggregations ---

    [Fact]
    public void Sum_SumsAllCells() => Assert.Equal(6.0, (double)MakeSingleColumn().Sum());

    [Fact]
    public void Min_ReturnsMinCell()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        Assert.Equal(1.0, (double)arr.Min());
    }

    [Fact]
    public void Max_ReturnsMaxCell()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        Assert.Equal(3.0, (double)arr.Max());
    }

    [Fact]
    public void Average_ReturnsAverageOfAllCells()
    {
        var arr = new ExcelArray(new object?[,] { { 2.0 }, { 4.0 }, { 6.0 } });
        Assert.Equal(4.0, (double)arr.Average());
    }

    [Fact]
    public void Aggregate_FoldsOverCells()
    {
        var result = MakeSingleColumn().Aggregate(
            new ExcelScalar(0.0),
            (acc, cell) => new ExcelScalar((double)acc + (double)cell));
        Assert.Equal(6.0, (double)result);
    }

    [Fact]
    public void Aggregate_DynamicSeed_And_DoubleReturns()
    {
        var result = MakeSingleColumn().Aggregate(0.0, (dynamic acc, dynamic cell) => acc + cell);
        Assert.Equal(6.0, (double)result);
    }

    [Fact]
    public void Aggregate_StringSeed_And_StringReturns()
    {
        var arr = new ExcelArray(new object?[,] { { "a" }, { "b" }, { "c" } });
        var result = arr.Aggregate("", (dynamic acc, dynamic cell) => (string)acc + (string)cell);
        Assert.Equal("abc", (string)result);
    }

    [Fact]
    public void Scan_RunningFold()
    {
        var result = MakeSingleColumn().Scan(
            new ExcelScalar(0.0),
            (acc, cell) => new ExcelScalar((double)acc + (double)cell));
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[1][0]);
        Assert.Equal(6.0, (double)rows[2][0]);
    }

    [Fact]
    public void Scan_DynamicSeed_And_DoubleReturns()
    {
        var result = MakeSingleColumn().Scan(0.0, (dynamic acc, dynamic cell) => acc + cell);
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[1][0]);
        Assert.Equal(6.0, (double)rows[2][0]);
    }

    // --- SelectMany ---

    [Fact]
    public void SelectMany_FlattensResults()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 2.0 } });
        var result = arr.SelectMany(v =>
            new ExcelValue[] { new ExcelScalar((double)v), new ExcelScalar((double)v * 10) });
        Assert.Equal(4, result.Count());
    }

    // --- Empty array edge cases ---

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

    // --- RowCollection-based operations (row-by-row) ---

    [Fact]
    public void Rows_Where_FiltersRowByRow()
    {
        var arr = MakeArray();
        var filtered = arr.Rows.Where(r => (double)r[0] > 1.0);
        Assert.Equal(2, filtered.Count());
    }

    [Fact]
    public void Rows_Any_TestsRowByRow()
    {
        var arr = MakeArray();
        Assert.True(arr.Rows.Any(r => (string)r[1] == "Bob"));
        Assert.False(arr.Rows.Any(r => (string)r[1] == "Dave"));
    }

    [Fact]
    public void Rows_OrderBy_SortsRowByRow()
    {
        var arr = new ExcelArray(new object?[,] { { 3.0 }, { 1.0 }, { 2.0 } });
        var sorted = arr.Rows.OrderBy(r => (double)r[0]);
        var rows = sorted.ToList();
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[2][0]);
    }

    [Fact]
    public void Rows_First_ReturnsFirstMatchingRow()
    {
        var arr = MakeArray();
        var row = arr.Rows.First(r => (double)r[0] > 1.0);
        Assert.Equal(2.0, (double)row[0]);
        Assert.Equal("Bob", row[1]);
    }

    [Fact]
    public void Rows_ToRange_ConvertsBackToExcelArray()
    {
        var arr = MakeArray();
        var filtered = arr.Rows.Where(r => (double)r[0] > 1.0);
        var range = filtered.ToRange();
        Assert.IsType<ExcelArray>(range);
        Assert.Equal(2, range.Rows.Count());
    }

    [Fact]
    public void Rows_All_TrueWhenAllRowsMatch()
    {
        var arr = MakeSingleColumn();
        Assert.True(arr.Rows.All(r => (double)r[0] > 0));
        Assert.False(arr.Rows.All(r => (double)r[0] > 2));
    }

    [Fact]
    public void Rows_FirstOrDefault_ReturnsNullWhenNoMatch()
    {
        var arr = MakeArray();
        var result = arr.Rows.FirstOrDefault(r => (double)r[0] > 100);
        Assert.Null(result);
    }

    [Fact]
    public void Rows_FirstOrDefault_ReturnsMatchingRow()
    {
        var arr = MakeArray();
        var result = arr.Rows.FirstOrDefault(r => (double)r[0] > 1.0);
        Assert.NotNull(result);
        Assert.Equal(2.0, (double)result[0]);
    }

    [Fact]
    public void Rows_OrderByDescending_SortsDescending()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0 }, { 3.0 }, { 2.0 } });
        var sorted = arr.Rows.OrderByDescending(r => (double)r[0]);
        var rows = sorted.ToList();
        Assert.Equal(3.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
        Assert.Equal(1.0, (double)rows[2][0]);
    }

    [Fact]
    public void Rows_Take_Positive_TakesFromStart()
    {
        var arr = MakeSingleColumn();
        var taken = arr.Rows.Take(2);
        Assert.Equal(2, taken.Count());
        var rows = taken.ToList();
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
    }

    [Fact]
    public void Rows_Take_Negative_TakesFromEnd()
    {
        var arr = MakeSingleColumn();
        var taken = arr.Rows.Take(-2);
        Assert.Equal(2, taken.Count());
        var rows = taken.ToList();
        Assert.Equal(2.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[1][0]);
    }

    [Fact]
    public void Rows_Skip_Positive_SkipsFromStart()
    {
        var arr = MakeSingleColumn();
        var skipped = arr.Rows.Skip(2);
        Assert.Equal(1, skipped.Count());
        Assert.Equal(3.0, (double)skipped.ToList()[0][0]);
    }

    [Fact]
    public void Rows_Skip_Negative_SkipsFromEnd()
    {
        var arr = MakeSingleColumn();
        var skipped = arr.Rows.Skip(-1);
        Assert.Equal(2, skipped.Count());
        var rows = skipped.ToList();
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(2.0, (double)rows[1][0]);
    }

    [Fact]
    public void Rows_Distinct_RemovesDuplicateRows()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0, "A" }, { 2.0, "B" }, { 1.0, "A" } });
        var distinct = arr.Rows.Distinct();
        Assert.Equal(2, distinct.Count());
    }

    // --- ExcelArray.Map ---

    [Fact]
    public void Map_Preserves2DShape()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
        var result = arr.Map(x => new ExcelScalar((double)x * 10));
        var resultArr = (object?[,])((ExcelValue)result).RawValue!;
        Assert.Equal(2, resultArr.GetLength(0));
        Assert.Equal(2, resultArr.GetLength(1));
        Assert.Equal(10.0, resultArr[0, 0]);
        Assert.Equal(20.0, resultArr[0, 1]);
        Assert.Equal(30.0, resultArr[1, 0]);
        Assert.Equal(40.0, resultArr[1, 1]);
    }
}
