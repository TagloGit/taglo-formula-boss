using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class CodeEmitterTests
{
    private readonly InputDetector _detector = new();
    private readonly CodeEmitter _emitter = new();

    private TranspileResult EmitFromExpression(string expression, string? preferredName = null)
    {
        var detection = _detector.Detect(expression);
        return _emitter.Emit(detection, expression, preferredName);
    }

    #region Using Directives

    [Fact]
    public void Emit_ContainsRuntimeUsing()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("using FormulaBoss.Runtime;", result.SourceCode);
    }

    [Fact]
    public void Emit_ContainsSystemLinqUsing()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("using System.Linq;", result.SourceCode);
    }

    #endregion

    #region Method Signature

    [Fact]
    public void Emit_MethodHasObjectParameters()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("object tbl__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_MultiInput_AllParametersPresent()
    {
        var result = EmitFromExpression("(tbl, maxVal) => tbl.Rows.Where(r => (double)r[0] < (double)maxVal)");

        Assert.Contains("object tbl__raw", result.SourceCode);
        Assert.Contains("object maxVal__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_ReturnTypeIsObject()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("public static object __udf_", result.SourceCode);
    }

    #endregion

    #region UDF Naming

    [Fact]
    public void Emit_UdfNameHasPrefix()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.StartsWith("__udf_", result.MethodName);
    }

    [Fact]
    public void Emit_SameExpression_SameHash()
    {
        var result1 = EmitFromExpression("tbl.Rows.Count()");
        var result2 = EmitFromExpression("tbl.Rows.Count()");

        Assert.Equal(result1.MethodName, result2.MethodName);
    }

    [Fact]
    public void Emit_DifferentExpressions_DifferentHash()
    {
        var result1 = EmitFromExpression("tbl.Rows.Count()");
        var result2 = EmitFromExpression("tbl.Rows.Sum()");

        Assert.NotEqual(result1.MethodName, result2.MethodName);
    }

    [Fact]
    public void Emit_PreferredName_Used()
    {
        var result = EmitFromExpression("tbl.Rows.Count()", "myCustomUdf");

        Assert.Equal("__udf_MYCUSTOMUDF", result.MethodName);
    }

    [Fact]
    public void Emit_ReservedExcelName_Prefixed()
    {
        // "SUM" is a reserved Excel name
        var result = EmitFromExpression("tbl.Rows.Count()", "sum");

        Assert.Equal("__udf__SUM", result.MethodName);
    }

    #endregion

    #region Input Wrapping

    [Fact]
    public void Emit_ContainsExcelReferenceCheck()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("tbl__isRef", result.SourceCode);
        Assert.Contains("\"ExcelReference\"", result.SourceCode);
    }

    [Fact]
    public void Emit_ContainsGetValuesFromReference()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("RuntimeHelpers.GetValuesFromReference", result.SourceCode);
    }

    [Fact]
    public void Emit_ContainsExcelValueWrap()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("ExcelValue.Wrap(", result.SourceCode);
    }

    [Fact]
    public void Emit_ContainsGetHeadersDelegate()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("GetHeadersDelegate", result.SourceCode);
    }

    #endregion

    #region Expression Body

    [Fact]
    public void Emit_SugarSyntax_ExpressionPassedThrough()
    {
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] > 5)");

        Assert.Contains("tbl.Rows.Where(r => (double)r[0] > 5)", result.SourceCode);
    }

    [Fact]
    public void Emit_ExplicitLambda_BodyExtracted()
    {
        var result = EmitFromExpression("(tbl) => tbl.Rows.Count()");

        Assert.Contains("var __result = tbl.Rows.Count()", result.SourceCode);
    }

    [Fact]
    public void Emit_StatementBody_WrappedInLocalFunction()
    {
        var result = EmitFromExpression("(tbl) => { var c = tbl.Rows.Count(); return c; }");

        Assert.Contains("__Execute()", result.SourceCode);
        Assert.Contains("var c = tbl.Rows.Count(); return c;", result.SourceCode);
    }

    #endregion

    #region Result Conversion

    [Fact]
    public void Emit_ContainsToResultDelegate()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("ToResultDelegate", result.SourceCode);
    }

    #endregion

    #region Object Model Propagation

    [Fact]
    public void Emit_CellAccess_RequiresObjectModel()
    {
        var result = EmitFromExpression("tbl.Cells.Where(c => c.Color == 6)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Emit_NoCellAccess_NoObjectModel()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.False(result.RequiresObjectModel);
    }

    [Fact]
    public void Emit_ObjectModel_IncludesOrigin()
    {
        var result = EmitFromExpression("tbl.Cells.Where(c => c.Color == 6)");

        Assert.Contains("GetOriginDelegate", result.SourceCode);
    }

    #endregion

    #region Free Variables

    [Fact]
    public void Emit_FreeVariable_BecomesParameter()
    {
        var result = EmitFromExpression("(tbl) => tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("object threshold__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_FreeVariable_WrappedAsScalar()
    {
        var result = EmitFromExpression("(tbl) => tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("ExcelValue.Wrap(threshold__raw)", result.SourceCode);
    }

    #endregion

    #region Class Structure

    [Fact]
    public void Emit_NoNamespace()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.DoesNotContain("namespace ", result.SourceCode);
    }

    [Fact]
    public void Emit_ClassNameMatchesMethod()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains($"{result.MethodName}_Class", result.SourceCode);
    }

    #endregion

    #region GenerateMethodName

    [Fact]
    public void GenerateMethodName_SanitizesInvalidChars()
    {
        var name = CodeEmitter.GenerateMethodName("x", "hello world!");

        Assert.Equal("__udf_HELLOWORLD", name);
    }

    [Fact]
    public void GenerateMethodName_PrependUnderscoreForDigitStart()
    {
        var name = CodeEmitter.GenerateMethodName("x", "123abc");

        Assert.Equal("__udf__123ABC", name);
    }

    [Fact]
    public void GenerateMethodName_EmptyPreferredName_UsesHash()
    {
        var name = CodeEmitter.GenerateMethodName("tbl.Rows.Count()", "   ");

        Assert.StartsWith("__udf_", name);
        Assert.NotEqual("__udf_", name);
    }

    #endregion
}
