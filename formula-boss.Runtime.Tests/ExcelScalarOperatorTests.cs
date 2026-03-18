using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class ExcelScalarOperatorTests
{
    [Fact]
    public void ImplicitDouble_ConvertsNumericValue()
    {
        var sv = new ExcelScalar(42.5);
        double d = sv;
        Assert.Equal(42.5, d);
    }

    [Fact]
    public void ImplicitString_ConvertsToString()
    {
        var sv = new ExcelScalar("hello");
        string? s = sv;
        Assert.Equal("hello", s);
    }

    [Fact]
    public void ImplicitBool_ConvertsBoolValue()
    {
        var sv = new ExcelScalar(true);
        bool b = sv;
        Assert.True(b);
    }

    [Fact]
    public void ComparisonOperators_WorkWithDoubles()
    {
        var sv = new ExcelScalar(10.0);
        Assert.True(sv > 5.0);
        Assert.True(sv < 15.0);
        Assert.True(sv >= 10.0);
        Assert.True(sv <= 10.0);
    }

    [Fact]
    public void ComparisonOperators_WorkBetweenExcelScalars()
    {
        var a = new ExcelScalar(10.0);
        var b = new ExcelScalar(20.0);
        Assert.True(a < b);
        Assert.True(b > a);
        Assert.False(a > b);
    }

    [Fact]
    public void ArithmeticOperators_WorkBetweenExcelScalars()
    {
        var a = new ExcelScalar(10.0);
        var b = new ExcelScalar(3.0);

        var sum = a + b;
        var diff = a - b;
        var prod = a * b;
        var quot = a / b;

        Assert.Equal(13.0, (double)sum);
        Assert.Equal(7.0, (double)diff);
        Assert.Equal(30.0, (double)prod);
        Assert.Equal(10.0 / 3.0, (double)quot);
    }

    [Fact]
    public void ArithmeticOperators_WorkWithDoubles()
    {
        var sv = new ExcelScalar(10.0);

        Assert.Equal(15.0, (double)(sv + 5.0));
        Assert.Equal(15.0, (double)(5.0 + sv));
        Assert.Equal(50.0, (double)(sv * 5.0));
    }

    [Fact]
    public void Equality_ComparesByValue()
    {
        var a = new ExcelScalar("test");
        var b = new ExcelScalar("test");
        var c = new ExcelScalar("other");

        Assert.True(a == b);
        Assert.False(a == c);
        Assert.True(a != c);
    }

    [Fact]
    public void Equality_WorksWithRawObjects()
    {
        var sv = new ExcelScalar("hello");
        Assert.True(sv == "hello");
        Assert.False(sv == "world");
    }

    [Fact]
    public void ToString_ReturnsStringRepresentation()
    {
        Assert.Equal("42", new ExcelScalar(42).ToString());
        Assert.Equal("hello", new ExcelScalar("hello").ToString());
        Assert.Equal("", new ExcelScalar(null).ToString());
    }

    [Fact]
    public void CrossTypeComparison_ExcelScalarVsExcelScalar()
    {
        var a = new ExcelScalar(20.0);
        var b = new ExcelScalar(10.0);

        Assert.True(a > b);
        Assert.False(a < b);
        Assert.True(a >= b);
        Assert.False(a <= b);
    }

    [Fact]
    public void CrossTypeEquality_ExcelScalarVsExcelValue()
    {
        var a = new ExcelScalar("hello");
        ExcelValue b = new ExcelScalar("hello");

        Assert.True(a == b);
        Assert.False(a != b);
        Assert.False(a == new ExcelScalar("other"));
    }
}
