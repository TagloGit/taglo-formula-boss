using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelScalarTests
{
    [Fact]
    public void RawValue_ReturnsWrappedValue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, scalar.RawValue);
    }

    [Fact]
    public void Rows_YieldsSingleRow()
    {
        var scalar = new ExcelScalar(42.0);
        var rows = scalar.Rows.ToList();
        Assert.Single(rows);
        Assert.Equal(42.0, (double)rows[0][0]);
    }

    [Fact]
    public void Where_MatchingPredicate_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.Where(_ => true);
        Assert.Equal(1, result.Count());
    }

    [Fact]
    public void Where_NonMatchingPredicate_ReturnsEmpty()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.Where(_ => false);
        Assert.Equal(0, result.Count());
    }

    [Fact]
    public void Any_MatchingPredicate_ReturnsTrue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.True(scalar.Any(_ => true));
    }

    [Fact]
    public void Any_NonMatchingPredicate_ReturnsFalse()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.False(scalar.Any(_ => false));
    }

    [Fact]
    public void All_MatchingPredicate_ReturnsTrue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.True(scalar.All(_ => true));
    }

    [Fact]
    public void Count_ReturnsOne()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(1, scalar.Count());
    }

    [Fact]
    public void Sum_ReturnsValue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar.Sum());
    }

    [Fact]
    public void Average_ReturnsValue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar.Average());
    }

    [Fact]
    public void First_MatchingPredicate_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.First(_ => true);
        Assert.Equal(42.0, (double)result);
    }

    [Fact]
    public void First_NonMatchingPredicate_Throws()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Throws<InvalidOperationException>(() => scalar.First(_ => false));
    }

    [Fact]
    public void FirstOrDefault_NonMatchingPredicate_ReturnsNull()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Null(scalar.FirstOrDefault(_ => false));
    }

    [Fact]
    public void Take_Zero_ReturnsEmpty()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(0, scalar.Take(0).Count());
    }

    [Fact]
    public void Skip_Zero_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(1, scalar.Skip(0).Count());
    }

    [Fact]
    public void Skip_One_ReturnsEmpty()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(0, scalar.Skip(1).Count());
    }

    [Fact]
    public void ImplicitConversions_Work()
    {
        var scalar = new ExcelScalar(42.0);
        double d = scalar;
        Assert.Equal(42.0, d);

        var strScalar = new ExcelScalar("hello");
        string? s = strScalar;
        Assert.Equal("hello", s);
    }

    [Fact]
    public void ComparisonOperators_Work()
    {
        var a = new ExcelScalar(10.0);
        var b = new ExcelScalar(20.0);
        Assert.True(a < b);
        Assert.True(b > a);
        Assert.True(a <= b);
        Assert.True(b >= a);
    }

    [Fact]
    public void Aggregate_AppliesFunction()
    {
        var scalar = new ExcelScalar(5.0);
        var result = scalar.Aggregate(
            new ExcelScalar(10.0),
            (acc, cell) => new ExcelScalar((double)acc + (double)cell));
        Assert.Equal(15.0, (double)result);
    }

    [Fact]
    public void Aggregate_DynamicSeed_And_DoubleReturns()
    {
        var scalar = new ExcelScalar(5.0);
        var result = scalar.Aggregate(10.0, (acc, cell) => acc + cell);
        Assert.Equal(15.0, (double)result);
    }

    [Fact]
    public void Scan_ReturnsResultOfFold()
    {
        var scalar = new ExcelScalar(5.0);
        var result = scalar.Scan(
            new ExcelScalar(10.0),
            (acc, cell) => new ExcelScalar((double)acc + (double)cell));
        Assert.Equal(15.0, (double)(ExcelValue)result);
    }

    [Fact]
    public void Scan_DynamicSeed_And_DoubleReturns()
    {
        var scalar = new ExcelScalar(5.0);
        var result = scalar.Scan(10.0, (acc, cell) => acc + cell);
        Assert.Equal(15.0, (double)(ExcelValue)result);
    }

    [Fact]
    public void Foreach_IteratesOnce()
    {
        var scalar = new ExcelScalar(42.0);
        var count = 0;
        var value = 0.0;
        foreach (var el in scalar)
        {
            value = (double)el;
            count++;
        }

        Assert.Equal(1, count);
        Assert.Equal(42.0, value);
    }

    [Fact]
    public void SelectMany_FlattensResults()
    {
        var scalar = new ExcelScalar(3.0);
        var result = scalar.SelectMany(v =>
            new ExcelValue[] { new ExcelScalar((double)v), new ExcelScalar((double)v * 10) });
        Assert.Equal(2, result.Count());
    }

    [Fact]
    public void Map_ReturnsMappedValue()
    {
        var scalar = new ExcelScalar(5.0);
        var result = scalar.Map(v => new ExcelScalar((double)v * 2));
        Assert.Equal(10.0, (double)(ExcelValue)result);
    }

    [Fact]
    public void OrderBy_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.OrderBy(v => v);
        Assert.Equal(42.0, (double)(ExcelValue)result);
    }

    [Fact]
    public void OrderByDescending_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.OrderByDescending(v => v);
        Assert.Equal(42.0, (double)(ExcelValue)result);
    }

    [Fact]
    public void Distinct_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        var result = scalar.Distinct();
        Assert.Equal(1, result.Count());
    }

    [Fact]
    public void Min_ReturnsValue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar.Min());
    }

    [Fact]
    public void Max_ReturnsValue()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar.Max());
    }

    // --- Indexers ---

    [Fact]
    public void Indexer2D_ZeroZero_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar[0, 0]);
    }

    [Fact]
    public void Indexer2D_NonZero_Throws()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Throws<IndexOutOfRangeException>(() => scalar[0, 1]);
        Assert.Throws<IndexOutOfRangeException>(() => scalar[1, 0]);
    }

    [Fact]
    public void IndexerLinear_Zero_ReturnsSelf()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(42.0, (double)scalar[0]);
    }

    [Fact]
    public void IndexerLinear_NonZero_Throws()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Throws<IndexOutOfRangeException>(() => scalar[1]);
        Assert.Throws<IndexOutOfRangeException>(() => scalar[-1]);
    }

    // --- IndexOf ---

    [Fact]
    public void IndexOf_MatchingValue_ReturnsZero()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(0, scalar.IndexOf(new ExcelScalar(42.0)));
    }

    [Fact]
    public void IndexOf_NonMatchingValue_ReturnsMinusOne()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(-1, scalar.IndexOf(new ExcelScalar(99.0)));
    }

    [Fact]
    public void IndexOf_RawValue_MatchingReturnsZero()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(0, scalar.IndexOf(42.0));
    }

    [Fact]
    public void IndexOf_RawValue_NonMatchingReturnsMinusOne()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(-1, scalar.IndexOf(99.0));
    }

    // --- RowCount / ColCount ---

    [Fact]
    public void RowCount_ReturnsOne()
    {
        var scalar = new ExcelScalar(42.0);
        Assert.Equal(1, scalar.RowCount);
    }

    [Fact]
    public void ColCount_ReturnsOne()
    {
        var scalar = new ExcelScalar("hello");
        Assert.Equal(1, scalar.ColCount);
    }

    [Fact]
    public void ArithmeticOperators_Work()
    {
        var a = new ExcelScalar(10.0);
        var b = new ExcelScalar(3.0);

        Assert.Equal(13.0, (double)(a + b));
        Assert.Equal(7.0, (double)(a - b));
        Assert.Equal(30.0, (double)(a * b));

        var c = new ExcelScalar(5.0);
        Assert.Equal(15.0, (double)(c * 3));
        Assert.Equal(7.0, (double)(c + 2));
    }
}
