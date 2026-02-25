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
            (acc, row) => new ExcelScalar((double)acc + (double)row[0]));
        Assert.Equal(15.0, (double)result);
    }
}
