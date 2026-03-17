using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelArrayTests
{
    private static ExcelArray MakeArray() =>
        new(new object?[,] { { 1.0, "Alice" }, { 2.0, "Bob" }, { 3.0, "Charlie" } });

    private static ExcelArray MakeSingleColumn() =>
        new(new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });

    // --- RowCount / ColCount ---

    [Fact]
    public void RowCount_ReturnsNumberOfRows()
    {
        var arr = MakeArray();
        Assert.Equal(3, arr.RowCount);
    }

    [Fact]
    public void ColCount_ReturnsNumberOfColumns()
    {
        var arr = MakeArray();
        Assert.Equal(2, arr.ColCount);
    }

    [Fact]
    public void RowCount_ColCount_AccessibleViaBaseType()
    {
        ExcelValue value = MakeArray();
        Assert.Equal(3, value.RowCount);
        Assert.Equal(2, value.ColCount);
    }

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
        var result = MakeSingleColumn().Aggregate(0.0, (acc, cell) => acc + cell);
        Assert.Equal(6.0, (double)result);
    }

    [Fact]
    public void Aggregate_StringSeed_And_StringReturns()
    {
        var arr = new ExcelArray(new object?[,] { { "a" }, { "b" }, { "c" } });
        var result = arr.Aggregate("", (acc, cell) => (string)acc + (string)cell);
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
        var result = MakeSingleColumn().Scan(0.0, (acc, cell) => acc + cell);
        var rows = result.Rows.ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(1.0, (double)rows[0][0]);
        Assert.Equal(3.0, (double)rows[1][0]);
        Assert.Equal(6.0, (double)rows[2][0]);
    }

    // --- Foreach (IEnumerable<ExcelValue>) ---

    [Fact]
    public void Foreach_IteratesElementWise()
    {
        var arr = MakeSingleColumn();
        var sum = 0.0;
        foreach (var el in arr)
        {
            sum += (double)el;
        }

        Assert.Equal(6.0, sum);
    }

    [Fact]
    public void Foreach_IteratesRowMajor_MultiColumn()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
        var values = new List<double>();
        foreach (var el in arr)
        {
            values.Add((double)el);
        }

        Assert.Equal(new[] { 1.0, 2.0, 3.0, 4.0 }, values);
    }

    [Fact]
    public void LinqToList_WorksOnArray()
    {
        var arr = MakeSingleColumn();
        var list = arr.ToList();
        Assert.Equal(3, list.Count);
        Assert.Equal(1.0, (double)list[0]);
    }

    [Fact]
    public void ForEach_ExecutesActionPerElement()
    {
        var arr = MakeSingleColumn();
        var sum = 0.0;
        arr.ForEach(el => sum += (double)el);
        Assert.Equal(6.0, sum);
    }

    [Fact]
    public void ForEach_Indexed_ProvidesCorrectRowAndColIndices()
    {
        var arr = new ExcelArray(new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 }, { 5.0, 6.0 } });
        var visited = new List<(double val, int row, int col)>();
        arr.ForEach((val, row, col) => visited.Add(((double)val, row, col)));

        Assert.Equal(6, visited.Count);
        Assert.Equal((1.0, 0, 0), visited[0]);
        Assert.Equal((2.0, 0, 1), visited[1]);
        Assert.Equal((3.0, 1, 0), visited[2]);
        Assert.Equal((4.0, 1, 1), visited[3]);
        Assert.Equal((5.0, 2, 0), visited[4]);
        Assert.Equal((6.0, 2, 1), visited[5]);
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

    // --- Indexers ---

    [Fact]
    public void Indexer2D_ReturnsCorrectElement()
    {
        var arr = MakeArray();
        Assert.Equal(1.0, (double)arr[0, 0]);
        Assert.Equal("Alice", arr[0, 1]);
        Assert.Equal(3.0, (double)arr[2, 0]);
        Assert.Equal("Charlie", arr[2, 1]);
    }

    [Fact]
    public void Indexer2D_ThrowsOnOutOfRange()
    {
        var arr = MakeArray();
        Assert.Throws<IndexOutOfRangeException>(() => arr[3, 0]);
        Assert.Throws<IndexOutOfRangeException>(() => arr[0, 2]);
        Assert.Throws<IndexOutOfRangeException>(() => arr[-1, 0]);
    }

    [Fact]
    public void IndexerLinear_ReturnsRowMajorOrder()
    {
        var arr = MakeArray(); // 3x2: {1,"Alice"},{2,"Bob"},{3,"Charlie"}
        Assert.Equal(1.0, (double)arr[0]);
        Assert.Equal("Alice", arr[1]);
        Assert.Equal(2.0, (double)arr[2]);
        Assert.Equal("Bob", arr[3]);
        Assert.Equal(3.0, (double)arr[4]);
        Assert.Equal("Charlie", arr[5]);
    }

    [Fact]
    public void IndexerLinear_ThrowsOnOutOfRange()
    {
        var arr = MakeArray();
        Assert.Throws<IndexOutOfRangeException>(() => arr[6]);
        Assert.Throws<IndexOutOfRangeException>(() => arr[-1]);
    }

    // --- IndexOf ---

    [Fact]
    public void IndexOf_FindsElementRowMajor()
    {
        var arr = MakeArray(); // 3x2: {1,"Alice"},{2,"Bob"},{3,"Charlie"}
        Assert.Equal(0, arr.IndexOf(new ExcelScalar(1.0)));
        Assert.Equal(1, arr.IndexOf(new ExcelScalar("Alice")));
        Assert.Equal(3, arr.IndexOf(new ExcelScalar("Bob")));
        Assert.Equal(5, arr.IndexOf(new ExcelScalar("Charlie")));
    }

    [Fact]
    public void IndexOf_ReturnsMinusOneWhenNotFound()
    {
        var arr = MakeArray();
        Assert.Equal(-1, arr.IndexOf(new ExcelScalar("Nobody")));
    }

    [Fact]
    public void IndexOf_WithIndexer_Roundtrips()
    {
        var arr = MakeArray();
        var idx = arr.IndexOf(new ExcelScalar("Bob"));
        Assert.Equal("Bob", arr[idx]);
    }

    [Fact]
    public void IndexOf_RawValue_FindsString()
    {
        var arr = MakeArray();
        Assert.Equal(1, arr.IndexOf("Alice"));
        Assert.Equal(3, arr.IndexOf("Bob"));
        Assert.Equal(-1, arr.IndexOf("Nobody"));
    }

    [Fact]
    public void IndexOf_RawValue_FindsDouble()
    {
        var arr = MakeArray();
        Assert.Equal(0, arr.IndexOf(1.0));
        Assert.Equal(2, arr.IndexOf(2.0));
        Assert.Equal(-1, arr.IndexOf(99.0));
    }

    [Fact]
    public void IndexOf_RawValue_WithIndexer_Roundtrips()
    {
        var arr = MakeArray();
        var idx = arr.IndexOf("Bob");
        Assert.Equal("Bob", arr[idx]);
    }

    // --- Map ---

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
