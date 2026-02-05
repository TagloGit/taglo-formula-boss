using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Integration tests for the value-only path (.values).
///     These tests compile generated code and run it with value arrays directly (no Excel COM needed).
/// </summary>
public class ValuePathTests
{
    private readonly ITestOutputHelper _output;

    public ValuePathTests(ITestOutputHelper output)
    {
        _output = output;
    }

    #region Complex Chains

    [Fact]
    public void Values_OrderByDesc_Take_Sum()
    {
        // Arrange: Top 3 values summed
        var values = new object[,] { { 10.0 }, { 50.0 }, { 30.0 }, { 20.0 }, { 40.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.orderByDesc(v => v).take(3).sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(120.0, result); // 50 + 40 + 30
    }

    #endregion

    #region Helpers

    private static string FormatResult(object? result)
    {
        if (result == null)
        {
            return "null";
        }

        if (result is string s)
        {
            return $"\"{s}\"";
        }

        if (result is object[,] arr)
        {
            var rows = arr.GetLength(0);
            var cols = arr.GetLength(1);
            var items = new List<string>();
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    items.Add(arr[r, c].ToString() ?? "null");
                }
            }

            return $"[{rows}x{cols}]: [{string.Join(", ", items)}]";
        }

        return result.ToString() ?? "null";
    }

    #endregion

    #region Basic Aggregations

    [Fact]
    public void Values_Sum_ReturnsCorrectTotal()
    {
        // Arrange
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.False(compilation.RequiresObjectModel);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(15.0, result);
    }

    [Fact]
    public void Values_Avg_ReturnsCorrectAverage()
    {
        // Arrange
        var values = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.avg()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(20.0, result);
    }

    [Fact]
    public void Values_Min_ReturnsMinimum()
    {
        // Arrange
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.min()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(1.0, result);
    }

    [Fact]
    public void Values_Max_ReturnsMaximum()
    {
        // Arrange
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.max()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(8.0, result);
    }

    [Fact]
    public void Values_Count_ReturnsCount()
    {
        // Arrange
        var values = new object[,] { { 1 }, { 2 }, { 3 }, { 4 }, { 5 } };
        var compilation = TestHelpers.CompileExpression("data.values.count()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(5, result);
    }

    #endregion

    #region Filtering

    [Fact]
    public void Values_WhereGreaterThan_ReturnsFilteredValues()
    {
        // Arrange
        var values = new object[,] { { 10.0 }, { 25.0 }, { 5.0 }, { 30.0 }, { 15.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.where(v => v > 20).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // 25 and 30
    }

    [Fact]
    public void Values_WhereWithSum_ReturnsSumOfFiltered()
    {
        // Arrange
        var values = new object[,] { { 10.0 }, { 25.0 }, { 5.0 }, { 30.0 }, { 15.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.where(v => v > 20).sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(55.0, result); // 25 + 30
    }

    #endregion

    #region Ordering

    [Fact]
    public void Values_OrderBy_SortsAscending()
    {
        // Arrange
        var values = new object[,] { { 3.0 }, { 1.0 }, { 4.0 }, { 1.0 }, { 5.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.orderBy(v => v).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(1.0, arr[0, 0]);
        Assert.Equal(1.0, arr[1, 0]);
        Assert.Equal(3.0, arr[2, 0]);
        Assert.Equal(4.0, arr[3, 0]);
        Assert.Equal(5.0, arr[4, 0]);
    }

    [Fact]
    public void Values_OrderByDesc_SortsDescending()
    {
        // Arrange
        var values = new object[,] { { 3.0 }, { 1.0 }, { 4.0 }, { 2.0 }, { 5.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.orderByDesc(v => v).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(5.0, arr[0, 0]);
        Assert.Equal(4.0, arr[1, 0]);
        Assert.Equal(3.0, arr[2, 0]);
        Assert.Equal(2.0, arr[3, 0]);
        Assert.Equal(1.0, arr[4, 0]);
    }

    #endregion

    #region Take and Skip

    [Fact]
    public void Values_Take_ReturnsFirstN()
    {
        // Arrange
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.take(3).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(3, arr.GetLength(0));
    }

    [Fact]
    public void Values_Skip_SkipsFirstN()
    {
        // Arrange
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 } };
        var compilation = TestHelpers.CompileExpression("data.values.skip(2).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");

        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(3.0, arr[0, 0]);
    }

    #endregion

    #region Named Column Access

    [Fact]
    public void Rows_WithHeaders_NamedColumnAccess_Where()
    {
        // Arrange: First row is headers
        var values = new object[,]
        {
            { "Name", "Price", "Qty" },
            { "Apple", 10.0, 5.0 },
            { "Banana", 20.0, 3.0 },
            { "Cherry", 15.0, 4.0 }
        };

        var compilation = TestHelpers.CompileExpression("data.withHeaders().rows.where(r => r[Price] > 12).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.False(compilation.RequiresObjectModel);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // Banana (20) and Cherry (15)
    }

    [Fact]
    public void Rows_WithHeaders_NamedColumnAccess_Reduce()
    {
        // Arrange: First row is headers
        var values = new object[,]
        {
            { "Name", "Price", "Qty" },
            { "Apple", 10.0, 2.0 },
            { "Banana", 20.0, 3.0 },
            { "Cherry", 15.0, 4.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r[Price] * r[Qty])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        // 10*2 + 20*3 + 15*4 = 20 + 60 + 60 = 140
        Assert.Equal(140.0, Convert.ToDouble(result));
    }

    [Fact]
    public void Rows_WithHeaders_DotNotation_Works()
    {
        // Arrange: First row is headers
        var values = new object[,]
        {
            { "Name", "Price", "Qty" },
            { "Apple", 10.0, 2.0 },
            { "Banana", 20.0, 3.0 }
        };

        // Use dot notation r.Price instead of bracket notation r[Price]
        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r.Price)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(30.0, Convert.ToDouble(result)); // 10 + 20
    }

    [Fact]
    public void Rows_WithHeaders_CaseInsensitive()
    {
        // Arrange: Headers in different case than access
        var values = new object[,]
        {
            { "NAME", "PRICE", "QTY" },
            { "Apple", 10.0, 2.0 },
            { "Banana", 20.0, 3.0 }
        };

        // Access with lowercase column names
        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r[price])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(30.0, Convert.ToDouble(result)); // 10 + 20
    }

    [Fact]
    public void Rows_WithHeaders_MissingColumn_ThrowsDetailedError()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price", "Qty" },
            { "Apple", 10.0, 2.0 }
        };

        // Access non-existent column
        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r[Cost])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert - should return error message with available columns
        _output.WriteLine($"Result: {result}");
        Assert.IsType<string>(result);
        var errorMsg = (string)result;
        Assert.Contains("Column 'Cost' not found", errorMsg);
        Assert.Contains("Name", errorMsg);
        Assert.Contains("Price", errorMsg);
        Assert.Contains("Qty", errorMsg);
    }

    [Fact]
    public void Rows_NumericIndex_StillWorks()
    {
        // Arrange: Use numeric index without headers
        var values = new object[,]
        {
            { 10.0, 2.0 },
            { 20.0, 3.0 },
            { 30.0, 4.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.rows.reduce(0, (acc, r) => acc + r[0] * r[1])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        // 10*2 + 20*3 + 30*4 = 20 + 60 + 120 = 200
        Assert.Equal(200.0, Convert.ToDouble(result));
    }

    #endregion
}
