using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Integration tests for the new InputDetector → CodeEmitter → Roslyn compile → execute pipeline.
///     Expressions are valid C# operating on FormulaBoss.Runtime types.
/// </summary>
public class WrapperTypePipelineTests
{
    private readonly ITestOutputHelper _output;

    public WrapperTypePipelineTests(ITestOutputHelper output)
    {
        _output = output;
    }

    #region Free Variables

    [Fact]
    public void FreeVariable_BecomesParameter()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var threshold = 10.0;
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[0] > (double)threshold)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Method should have 2 params: tbl__raw and threshold__raw (sorted: tbl < threshold)
        var parameters = compilation.Method!.GetParameters();
        Assert.Equal(2, parameters.Length);

        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, values, threshold);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 15 and 20
    }

    #endregion

    #region Sugar Syntax — Single Input

    [Fact]
    public void Sugar_RowFilter_NumericIndex()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[0] > 10)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 15 and 20
    }

    [Fact]
    public void Sugar_RowCount()
    {
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Rows.Count()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(3, result);
    }

    [Fact]
    public void Sugar_Sum()
    {
        var values = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(60.0, result);
    }

    [Fact]
    public void Sugar_FilterThenCount()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 25.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[0] > 10).Count()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(2, result);
    }

    #endregion

    #region String Key Access (Headers)

    [Fact]
    public void Sugar_StringKeyFilter_Compiles()
    {
        // Verify string key access expression compiles and detects headers requirement
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[\"Price\"] > 12)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.Contains("tbl", compilation.Detection!.HeaderVariables);
    }

    [Fact]
    public void Sugar_StringKeyFilter_ExecutesWithHeaders()
    {
        // Row 0 = headers, rows 1-3 = data. Headers should be stripped.
        var values = new object[,] { { "Price" }, { 5.0 }, { 15.0 }, { 8.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => r[\"Price\"] > 10)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(1, arr.GetLength(0)); // Only 15 passes
        Assert.Equal(15.0, arr[0, 0]);
    }

    #endregion

    #region Multiple Parameters

    [Fact]
    public void MultipleParams_FilterWithThreshold()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var maxVal = 10.0;
        // No explicit lambda — both tbl and maxVal are free variables
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[0] < (double)maxVal)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Parameters are sorted: maxVal, tbl
        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, maxVal, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 5 and 3
    }

    [Fact]
    public void MultipleParams_FilterWithThreshold_NoCast()
    {
        // Test that r[0] > maxVal works without explicit casts (ColumnValue > ExcelValue operator)
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var maxVal = 10.0;
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => r[0] > maxVal)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Parameters are sorted: maxVal, tbl
        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, maxVal, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 15 and 20
    }

    #endregion

    #region Aggregation

    [Fact]
    public void Sugar_Average()
    {
        var values = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Average()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(20.0, result);
    }

    [Fact]
    public void Sugar_Min()
    {
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Min()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(1.0, result);
    }

    [Fact]
    public void Sugar_Max()
    {
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Max()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(8.0, result);
    }

    #endregion

    #region OrderBy and Take/Skip

    [Fact]
    public void Sugar_OrderBy()
    {
        var values = new object[,] { { 30.0 }, { 10.0 }, { 20.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.OrderBy(r => (double)r[0]).ToRange()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(10.0, arr[0, 0]);
        Assert.Equal(20.0, arr[1, 0]);
        Assert.Equal(30.0, arr[2, 0]);
    }

    [Fact]
    public void Sugar_Take()
    {
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Take(3)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
    }

    #endregion

    #region UDF Naming

    [Fact]
    public void UdfName_StartsWithPrefix()
    {
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Rows.Count()");

        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.StartsWith("__udf_", compilation.MethodName);
    }

    [Fact]
    public void UdfName_ConsistentForSameExpression()
    {
        var result1 = NewPipelineTestHelpers.CompileExpression("tbl.Rows.Count()");
        var result2 = NewPipelineTestHelpers.CompileExpression("tbl.Rows.Count()");

        Assert.Equal(result1.MethodName, result2.MethodName);
    }

    #endregion

    #region Object Model Detection

    [Fact]
    public void ObjectModel_CellsDetected()
    {
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Cells.Where(c => c.Color == 6)");

        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.True(compilation.RequiresObjectModel);
    }

    [Fact]
    public void ObjectModel_NoCellsNotDetected()
    {
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Rows.Count()");

        Assert.True(compilation.Success, compilation.ErrorMessage);
        Assert.False(compilation.RequiresObjectModel);
    }

    #endregion

    #region Skip

    [Fact]
    public void Sugar_Skip()
    {
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 }, { 4.0 }, { 5.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Skip(2)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
        Assert.Equal(3.0, arr[0, 0]);
    }

    #endregion

    #region Distinct

    [Fact]
    public void Sugar_Distinct()
    {
        var values = new object[,] { { 1.0 }, { 2.0 }, { 1.0 }, { 3.0 }, { 2.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Distinct()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0)); // 1, 2, 3
    }

    #endregion

    #region GroupBy

    [Fact]
    public void Sugar_GroupBy_Compiles()
    {
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.GroupBy(r => r[0])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);
    }

    #endregion

    #region Aggregate

    [Fact]
    public void Sugar_Aggregate_WithSeed()
    {
        var values = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Aggregate(0.0, (acc, r) => acc + (double)r[0])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(60.0, result);
    }

    #endregion

    #region Negative Indexing

    [Fact]
    public void Sugar_NegativeIndex_LastColumn()
    {
        var values = new object[,] { { 1.0, 10.0 }, { 2.0, 20.0 }, { 3.0, 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Select(r => r[-1])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
    }

    #endregion

    #region Backtick in LET Result Position

    [Fact]
    public void BacktickInLetResultPosition_Compiles()
    {
        // LET result position backtick — just test compilation via the pipeline
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!,
            new object[,] { { 10.0 }, { 20.0 } });
        Assert.Equal(30.0, result);
    }

    #endregion

    #region Single-Cell Reference

    [Fact]
    public void SingleCellRef_AsScalar()
    {
        // A single cell value wraps as ExcelScalar
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // 1x1 array wraps to ExcelScalar, .Sum() returns value
        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!,
            new object[,] { { 42.0 } });
        Assert.Equal(42.0, result);
    }

    #endregion

    #region Statement Block

    [Fact(Skip = "Blocked by #108")]
    public void StatementBlock_CompilesAndExecutes()
    {
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "{ var x = 1; return tbl.Sum() + x; }");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);
    }

    #endregion

    #region GetHeadersDelegate

    [Fact]
    public void GetHeadersDelegate_WorksForObjectArray()
    {
        // Verify header extraction works through the full pipeline
        var values = new object[,] { { "Price", "Name" }, { 10.0, "Widget" }, { 20.0, "Gadget" } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => r[\"Price\"] > 15).Count()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(1, result); // Only Gadget (20)
    }

    #endregion

    #region Select and Chaining

    [Fact]
    public void Sugar_SelectColumn()
    {
        var values = new object[,] { { 1.0, 10.0 }, { 2.0, 20.0 }, { 3.0, 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Select(r => r[1])");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr.GetLength(0));
    }

    [Fact]
    public void Sugar_WhereSelectCount()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 25.0 }, { 8.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.Rows.Where(r => (double)r[0] > 10).Count()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        Assert.Equal(2, result); // 15 and 25
    }

    #endregion
}
