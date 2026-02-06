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

    [Fact]
    public void Rows_WithHeaders_StringComparison()
    {
        // Arrange: Filter by string column
        var values = new object[,]
        {
            { "Name", "Category", "Price" },
            { "Apple", "Fruit", 10.0 },
            { "Carrot", "Vegetable", 8.0 },
            { "Banana", "Fruit", 25.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.where(r => r[Category] == \"Fruit\").toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // Apple and Banana (Fruit rows)
        Assert.Equal("Apple", arr[0, 0]);
        Assert.Equal("Banana", arr[1, 0]);
    }

    [Fact]
    public void Rows_WithHeaders_StringComparison_ChainedWithReduce()
    {
        // Arrange: Filter by category, then sum prices
        var values = new object[,]
        {
            { "Name", "Category", "Price" },
            { "Apple", "Fruit", 10.0 },
            { "Carrot", "Vegetable", 8.0 },
            { "Banana", "Fruit", 25.0 },
            { "Date", "Fruit", 30.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.where(r => r[Category] == \"Fruit\").reduce(0, (acc, r) => acc + r[Price])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        // 10 + 25 + 30 = 65 (only Fruit prices)
        Assert.Equal(65.0, Convert.ToDouble(result));
    }

    [Fact]
    public void Rows_NegativeIndex_AccessesLastColumn()
    {
        // Arrange: r[-1] should access the last column
        var values = new object[,]
        {
            { "A", "B", "C", 100.0 },
            { "D", "E", "F", 200.0 },
            { "G", "H", "I", 300.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.rows.select(r => r[-1]).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        var array = (object[,])result;
        _output.WriteLine($"Result dimensions: {array.GetLength(0)}x{array.GetLength(1)}");
        Assert.Equal(100.0, Convert.ToDouble(array[0, 0]));
        Assert.Equal(200.0, Convert.ToDouble(array[1, 0]));
        Assert.Equal(300.0, Convert.ToDouble(array[2, 0]));
    }

    [Fact]
    public void Rows_NegativeIndex_SecondFromEnd()
    {
        // Arrange: r[-2] should access the second to last column
        var values = new object[,]
        {
            { "A", "B", "C", 100.0 },
            { "D", "E", "F", 200.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.rows.select(r => r[-2]).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        var array = (object[,])result;
        Assert.Equal("C", array[0, 0]);
        Assert.Equal("F", array[1, 0]);
    }

    #endregion

    #region Row Predicate Methods (find, some, every)

    [Fact]
    public void Rows_Find_ReturnsFirstMatchingRow()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 25.0 },
            { "Cherry", 15.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.find(r => r[Price] > 20)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        // Should return the Banana row - displayed as vertical column (N rows, 1 col)
        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0)); // 2 values (Name, Price values)
        Assert.Equal(1, arr.GetLength(1)); // Single column
        Assert.Equal("Banana", arr[0, 0]); // Name
        Assert.Equal(25.0, Convert.ToDouble(arr[1, 0])); // Price
    }

    [Fact]
    public void Rows_Find_ReturnsNullWhenNoMatch()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 20.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.find(r => r[Price] > 100)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        // Should return empty (no match)
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void Rows_Some_ReturnsTrueWhenMatchExists()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 25.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.some(r => r[Price] > 20)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(true, result);
    }

    [Fact]
    public void Rows_Some_ReturnsFalseWhenNoMatch()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 20.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.some(r => r[Price] > 100)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(false, result);
    }

    [Fact]
    public void Rows_Every_ReturnsTrueWhenAllMatch()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 20.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.every(r => r[Price] > 0)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(true, result);
    }

    [Fact]
    public void Rows_Every_ReturnsFalseWhenSomeDontMatch()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Price" },
            { "Apple", 10.0 },
            { "Banana", 20.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.every(r => r[Price] > 15)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(false, result);
    }

    #endregion

    #region Scan Method

    [Fact]
    public void Rows_Scan_ReturnsRunningTotals()
    {
        // Arrange
        var values = new object[,]
        {
            { "Name", "Amount" },
            { "A", 10.0 },
            { "B", 20.0 },
            { "C", 30.0 }
        };

        var compilation = TestHelpers.CompileExpression(
            "data.withHeaders().rows.scan(0, (sum, r) => sum + r[Amount]).toArray()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValues(compilation.CoreMethod!, values);

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.NotNull(result);
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        // Running totals: 10, 30, 60 (output as 3 rows, 1 column - vertical)
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(1, arr.GetLength(1));
        Assert.Equal(10.0, Convert.ToDouble(arr[0, 0]));
        Assert.Equal(30.0, Convert.ToDouble(arr[1, 0]));
        Assert.Equal(60.0, Convert.ToDouble(arr[2, 0]));
    }

    #endregion

    #region Dynamic Column Name Parameters (LET Column Bindings)

    [Fact]
    public void ColumnBindings_ReduceWithDynamicColumnNames_Works()
    {
        // Arrange: Simulate a LET formula with column bindings
        // =LET(tbl, tblSales, p, tblSales[Price], q, tblSales[Qty], `tbl.rows.reduce(0, (acc, r) => acc + r.p * r.q)`)
        var values = new object[,]
        {
            { "Name", "Price", "Qty" },
            { "Apple", 10.0, 2.0 },
            { "Banana", 20.0, 3.0 },
            { "Cherry", 15.0, 4.0 }
        };

        // Column bindings: p -> Price, q -> Qty
        var columnBindings = new Dictionary<string, FormulaBoss.Interception.ColumnBindingInfo>
        {
            ["p"] = new("tblSales", "Price"),
            ["q"] = new("tblSales", "Qty")
        };

        // Expression uses r.p and r.q which should resolve via column bindings
        var compilation = TestHelpers.CompileExpressionWithColumnBindings(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r.p * r.q)",
            columnBindings);

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Verify that column bindings were detected and used
        Assert.NotNull(compilation.UsedColumnBindings);
        Assert.Equal(2, compilation.UsedColumnBindings.Count);
        Assert.Contains("p", compilation.UsedColumnBindings);
        Assert.Contains("q", compilation.UsedColumnBindings);

        // Act: Call _Core with column names passed as parameters (simulating Excel passing INDEX() results)
        var result = TestHelpers.ExecuteWithValuesAndColumnNames(
            compilation.CoreMethod!,
            values,
            "Price", "Qty");  // These are the actual column names

        // Assert
        _output.WriteLine($"Result: {result}");
        // 10*2 + 20*3 + 15*4 = 20 + 60 + 60 = 140
        Assert.Equal(140.0, Convert.ToDouble(result));
    }

    [Fact]
    public void ColumnBindings_SelectWithDynamicColumnNames_Works()
    {
        // Arrange
        var values = new object[,]
        {
            { "Product", "UnitPrice", "Quantity" },
            { "A", 5.0, 3.0 },
            { "B", 10.0, 2.0 }
        };

        // Column bindings with different names than headers (common pattern)
        var columnBindings = new Dictionary<string, FormulaBoss.Interception.ColumnBindingInfo>
        {
            ["price"] = new("tbl", "UnitPrice"),
            ["qty"] = new("tbl", "Quantity")
        };

        var compilation = TestHelpers.CompileExpressionWithColumnBindings(
            "data.withHeaders().rows.select(r => r.price * r.qty)",
            columnBindings);

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act: Pass actual column names
        var result = TestHelpers.ExecuteWithValuesAndColumnNames(
            compilation.CoreMethod!,
            values,
            "UnitPrice", "Quantity");

        // Assert
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0));
        Assert.Equal(15.0, Convert.ToDouble(arr[0, 0])); // 5*3
        Assert.Equal(20.0, Convert.ToDouble(arr[1, 0])); // 10*2
    }

    [Fact]
    public void ColumnBindings_WhereWithDynamicColumnNames_Works()
    {
        // Arrange
        var values = new object[,]
        {
            { "Item", "Cost" },
            { "A", 50.0 },
            { "B", 150.0 },
            { "C", 75.0 }
        };

        var columnBindings = new Dictionary<string, FormulaBoss.Interception.ColumnBindingInfo>
        {
            ["c"] = new("data", "Cost")
        };

        var compilation = TestHelpers.CompileExpressionWithColumnBindings(
            "data.withHeaders().rows.where(r => r.c > 60)",
            columnBindings);

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValuesAndColumnNames(
            compilation.CoreMethod!,
            values,
            "Cost");

        // Assert - should return B (150) and C (75)
        _output.WriteLine($"Result: {FormatResult(result)}");
        Assert.IsType<object[,]>(result);
        var arr = (object[,])result;
        Assert.Equal(2, arr.GetLength(0));
    }

    [Fact]
    public void ColumnBindings_BracketNotation_Works()
    {
        // Test r[price] bracket notation with column bindings
        var values = new object[,]
        {
            { "Name", "Price" },
            { "X", 100.0 },
            { "Y", 200.0 }
        };

        var columnBindings = new Dictionary<string, FormulaBoss.Interception.ColumnBindingInfo>
        {
            ["price"] = new("tbl", "Price")
        };

        // Use bracket notation r[price] instead of r.price
        var compilation = TestHelpers.CompileExpressionWithColumnBindings(
            "data.withHeaders().rows.reduce(0, (acc, r) => acc + r[price])",
            columnBindings);

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Act
        var result = TestHelpers.ExecuteWithValuesAndColumnNames(
            compilation.CoreMethod!,
            values,
            "Price");

        // Assert
        _output.WriteLine($"Result: {result}");
        Assert.Equal(300.0, Convert.ToDouble(result)); // 100 + 200
    }

    #endregion
}
