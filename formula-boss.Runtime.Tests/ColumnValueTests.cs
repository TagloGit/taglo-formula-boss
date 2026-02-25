using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ColumnValueTests
{
    [Fact]
    public void ImplicitDouble_ConvertsNumericValue()
    {
        var cv = new ColumnValue(42.5);
        double d = cv;
        Assert.Equal(42.5, d);
    }

    [Fact]
    public void ImplicitString_ConvertsToString()
    {
        var cv = new ColumnValue("hello");
        string? s = cv;
        Assert.Equal("hello", s);
    }

    [Fact]
    public void ImplicitBool_ConvertsBoolValue()
    {
        var cv = new ColumnValue(true);
        bool b = cv;
        Assert.True(b);
    }

    [Fact]
    public void ComparisonOperators_WorkWithDoubles()
    {
        var cv = new ColumnValue(10.0);
        Assert.True(cv > 5.0);
        Assert.True(cv < 15.0);
        Assert.True(cv >= 10.0);
        Assert.True(cv <= 10.0);
    }

    [Fact]
    public void ComparisonOperators_WorkBetweenColumnValues()
    {
        var a = new ColumnValue(10.0);
        var b = new ColumnValue(20.0);
        Assert.True(a < b);
        Assert.True(b > a);
        Assert.False(a > b);
    }

    [Fact]
    public void ArithmeticOperators_WorkBetweenColumnValues()
    {
        var a = new ColumnValue(10.0);
        var b = new ColumnValue(3.0);

        ColumnValue sum = a + b;
        ColumnValue diff = a - b;
        ColumnValue prod = a * b;
        ColumnValue quot = a / b;

        Assert.Equal(13.0, (double)sum);
        Assert.Equal(7.0, (double)diff);
        Assert.Equal(30.0, (double)prod);
        Assert.Equal(10.0 / 3.0, (double)quot);
    }

    [Fact]
    public void ArithmeticOperators_WorkWithDoubles()
    {
        var cv = new ColumnValue(10.0);

        Assert.Equal(15.0, (double)(cv + 5.0));
        Assert.Equal(15.0, (double)(5.0 + cv));
        Assert.Equal(50.0, (double)(cv * 5.0));
    }

    [Fact]
    public void Equality_ComparesByValue()
    {
        var a = new ColumnValue("test");
        var b = new ColumnValue("test");
        var c = new ColumnValue("other");

        Assert.True(a == b);
        Assert.False(a == c);
        Assert.True(a != c);
    }

    [Fact]
    public void Equality_WorksWithRawObjects()
    {
        var cv = new ColumnValue("hello");
        Assert.True(cv == (object)"hello");
        Assert.False(cv == (object)"world");
    }

    [Fact]
    public void ToString_ReturnsStringRepresentation()
    {
        Assert.Equal("42", new ColumnValue(42).ToString());
        Assert.Equal("hello", new ColumnValue("hello").ToString());
        Assert.Equal("", new ColumnValue(null).ToString());
    }
}
