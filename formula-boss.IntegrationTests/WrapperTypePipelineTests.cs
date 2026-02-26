using System.Reflection;

using FormulaBoss.Runtime;
using FormulaBoss.Transpilation;

using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Integration tests that compile generated code and execute it against real wrapper types.
///     These tests validate the full pipeline: expression → InputDetector → CodeEmitter → Roslyn compile → execute.
/// </summary>
public class WrapperTypePipelineTests
{
    private readonly ITestOutputHelper _output;

    public WrapperTypePipelineTests(ITestOutputHelper output)
    {
        _output = output;
    }

    // === Single-input sugar syntax ===

    [Fact]
    public void Sugar_RowsWhere_FiltersRows()
    {
        // 3 rows, 2 cols: Name, Value
        var data = new object[,] { { "Alice", 10.0 }, { "Bob", 20.0 }, { "Carol", 5.0 } };

        var result = CompileAndExecute("tbl.Rows.Where(r => r[1] > 8)", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // Alice(10) and Bob(20)
    }

    [Fact]
    public void Sugar_RowsCount_ReturnsCount()
    {
        var data = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };

        var result = CompileAndExecute("tbl.Count()", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, Convert.ToInt32(arr[0, 0]));
    }

    [Fact]
    public void Sugar_RowsSum_ReturnsSumValue()
    {
        var data = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };

        var result = CompileAndExecute("tbl.Sum()", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(60.0, Convert.ToDouble(arr[0, 0]));
    }

    [Fact]
    public void Sugar_RowsAny_ReturnsBool()
    {
        var data = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };

        var result = CompileAndExecute("tbl.Any(r => r[0] > 2)", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(true, arr[0, 0]);
    }

    [Fact]
    public void Sugar_BracketAccess_WithStringKey()
    {
        // Simulate a table with headers by using ExcelTable directly
        var data = new object[,] { { 10.0, "USD" }, { 20.0, "EUR" }, { 5.0, "GBP" } };
        var headers = new[] { "Unit Price", "Currency" };

        var result = CompileAndExecuteWithHeaders(
            "tbl.Rows.Where(r => r[\"Unit Price\"] == 10)", data, headers);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(1, arr.GetLength(0));
        Assert.Equal(10.0, Convert.ToDouble(arr[0, 0]));
    }

    [Fact]
    public void Sugar_ChainedOperations_WhereAndSelect()
    {
        var data = new object[,] { { 10.0 }, { 20.0 }, { 5.0 } };

        var result = CompileAndExecute(
            "tbl.Rows.Where(r => r[0] > 8).Select(r => r[0])", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0));
    }

    [Fact]
    public void Sugar_OrderBy_SortsRows()
    {
        var data = new object[,] { { 30.0 }, { 10.0 }, { 20.0 } };

        var result = CompileAndExecute("tbl.OrderBy(r => (double)r[0])", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(10.0, Convert.ToDouble(arr[0, 0]));
        Assert.Equal(20.0, Convert.ToDouble(arr[1, 0]));
        Assert.Equal(30.0, Convert.ToDouble(arr[2, 0]));
    }

    // === Explicit lambda syntax ===

    [Fact]
    public void ExplicitLambda_SingleInput_Works()
    {
        var data = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };

        var result = CompileAndExecute("(tbl) => tbl.Count()", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, Convert.ToInt32(arr[0, 0]));
    }

    [Fact]
    public void ExplicitLambda_MultiInput_ScalarComparison()
    {
        var data = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var threshold = 15.0;

        var result = CompileAndExecuteMultiInput(
            "(tbl, maxVal) => tbl.Rows.Where(r => r[0] > (double)maxVal)",
            data, threshold);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 20 and 30
    }

    // === Statement block syntax ===

    [Fact]
    public void StatementBlock_WithReturn_Works()
    {
        var data = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };

        var result = CompileAndExecute(
            "(tbl) => { if (tbl.Count() > 2) return \"big\"; return \"small\"; }", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal("big", arr[0, 0]);
    }

    // === Cell/Cells detection ===

    [Fact]
    public void CellUsage_SetsRequiresObjectModel()
    {
        var expr = "tbl.Rows.Where(r => r[0].Cell.Color == 6)";
        var detection = InputDetector.Detect(expr);
        var result = CodeEmitter.Emit(detection, "test", expr);

        Assert.True(result.RequiresObjectModel);
    }

    // === Nested lambdas with free variables ===

    [Fact]
    public void NestedLambda_FreeVariable_Compiles()
    {
        // tbl.Rows.Where(r => pConts.Any(c => c[0] == r[1]))
        // where pConts is a secondary input from LET
        var tblData = new object[,] { { 10.0, "US" }, { 20.0, "UK" }, { 30.0, "FR" } };
        var conts = new object[,] { { "US" }, { "FR" } };

        var expr = "(tbl, pConts) => tbl.Rows.Where(r => pConts.Any(c => c[0] == r[1]))";
        var result = CompileAndExecuteMultiInput(expr, tblData, conts);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // US and FR rows
    }

    // === Empty results ===

    [Fact]
    public void Where_NoMatch_ReturnsEmptyArray()
    {
        var data = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };

        var result = CompileAndExecute("tbl.Rows.Where(r => r[0] > 100)", data);

        Assert.NotNull(result);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(0, arr.GetLength(0));
    }

    // === Helpers ===

    private object? CompileAndExecute(string expression, object[,] data)
    {
        var compiled = CompileExpression(expression);
        Assert.True(compiled.Success, compiled.GetDiagnostics());

        // Invoke with raw data (simulating what Excel passes after ExcelReference resolution)
        return compiled.CoreMethod!.Invoke(null, new object[] { data });
    }

    private object? CompileAndExecuteWithHeaders(string expression, object[,] data, string[] headers)
    {
        // Create an ExcelTable and pass it — the Wrap() call in generated code will pass it through
        var table = new ExcelTable(data, headers);
        var compiled = CompileExpression(expression);
        Assert.True(compiled.Success, compiled.GetDiagnostics());

        return compiled.CoreMethod!.Invoke(null, new object[] { table });
    }

    private object? CompileAndExecuteMultiInput(string expression, object[,] data, object secondInput)
    {
        var compiled = CompileExpression(expression);
        Assert.True(compiled.Success, compiled.GetDiagnostics());

        return compiled.CoreMethod!.Invoke(null, new object[] { data, secondInput });
    }

    private TestCompilationResult CompileExpression(string expression)
    {
        var detection = InputDetector.Detect(expression);
        var transpileResult = CodeEmitter.Emit(detection, expression, expression);

        _output.WriteLine("=== Generated Code ===");
        _output.WriteLine(transpileResult.SourceCode);
        _output.WriteLine("=== End ===");

        var result = TestHelpers.CompileExpression(expression);
        return result;
    }
}
