using FormulaBoss.Runtime;

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
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr[0, 0]);
    }

    [Fact]
    public void Sugar_Sum()
    {
        var values = new object[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Sum()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(60.0, arr[0, 0]);
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
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr[0, 0]);
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
        Assert.True(compilation.Detection!.HasStringBracketAccess);
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

    #region Explicit Lambda — Multi-Input

    [Fact]
    public void ExplicitLambda_MultiInput()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var maxVal = 10.0;
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "(tbl, maxVal) => tbl.Rows.Where(r => (double)r[0] < (double)maxVal)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, values, maxVal);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 5 and 3
    }

    [Fact]
    public void ExplicitLambda_MultiInput_NoCast()
    {
        // Test that r[0] > maxVal works without explicit casts (ColumnValue > ExcelValue operator)
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var maxVal = 10.0;
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "(tbl, maxVal) => tbl.Rows.Where(r => r[0] > maxVal)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, values, maxVal);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 15 and 20
    }

    #endregion

    #region Statement Block

    [Fact]
    public void StatementBlock_ReturnsCount()
    {
        var values = new object[,] { { 1.0 }, { 2.0 }, { 3.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "(tbl) => { var c = tbl.Rows.Count(); return c; }");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, arr[0, 0]);
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
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(20.0, arr[0, 0]);
    }

    [Fact]
    public void Sugar_Min()
    {
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Min()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(1.0, arr[0, 0]);
    }

    [Fact]
    public void Sugar_Max()
    {
        var values = new object[,] { { 5.0 }, { 2.0 }, { 8.0 }, { 1.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression("tbl.Max()");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        var result = NewPipelineTestHelpers.ExecuteWithValues(compilation.Method!, values);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(8.0, arr[0, 0]);
    }

    #endregion

    #region OrderBy and Take/Skip

    [Fact]
    public void Sugar_OrderBy()
    {
        var values = new object[,] { { 30.0 }, { 10.0 }, { 20.0 } };
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "tbl.OrderBy(r => (double)r[0])");

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

    #region Select and Chaining

    [Fact]
    public void Sugar_SelectColumn()
    {
        var values = new object[,]
        {
            { 1.0, 10.0 },
            { 2.0, 20.0 },
            { 3.0, 30.0 }
        };
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
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr[0, 0]); // 15 and 25
    }

    #endregion

    #region Free Variables

    [Fact]
    public void FreeVariable_BecomesParameter()
    {
        var values = new object[,] { { 5.0 }, { 15.0 }, { 3.0 }, { 20.0 } };
        var threshold = 10.0;
        var compilation = NewPipelineTestHelpers.CompileExpression(
            "(tbl) => tbl.Rows.Where(r => (double)r[0] > (double)threshold)");

        _output.WriteLine(compilation.GetDiagnostics());
        Assert.True(compilation.Success, compilation.ErrorMessage);

        // Method should have 2 params: tbl__raw and threshold__raw
        var parameters = compilation.Method!.GetParameters();
        Assert.Equal(2, parameters.Length);

        var result = NewPipelineTestHelpers.ExecuteWithMultipleInputs(compilation.Method!, values, threshold);
        var arr = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, arr.GetLength(0)); // 15 and 20
    }

    #endregion
}
