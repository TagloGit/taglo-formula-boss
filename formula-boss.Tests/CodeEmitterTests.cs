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

    #region Result Conversion

    [Fact]
    public void Emit_ContainsToResultDelegate()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains("ToResultDelegate", result.SourceCode);
    }

    #endregion

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
    public void Emit_MultipleParams_AllParametersPresent()
    {
        // Both tbl and maxVal are free variables, detected as flat parameters
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] < (double)maxVal)");

        Assert.Contains("object maxVal__raw", result.SourceCode);
        Assert.Contains("object tbl__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_ReturnTypeIsObject()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.Contains($"public static object {CodeEmitter.UdfPrefix}", result.SourceCode);
    }

    #endregion

    #region UDF Naming

    [Fact]
    public void Emit_UdfNameHasPrefix()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.StartsWith(CodeEmitter.UdfPrefix, result.MethodName);
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

        Assert.Equal($"{CodeEmitter.UdfPrefix}MYCUSTOMUDF", result.MethodName);
    }

    [Fact]
    public void Emit_ReservedExcelName_Prefixed()
    {
        var result = EmitFromExpression("tbl.Rows.Count()", "sum");

        Assert.Equal($"{CodeEmitter.UdfPrefix}_SUM", result.MethodName);
    }

    #endregion

    #region Parameter Wrapping

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
    public void Emit_ContainsGetHeadersDelegate_WhenStringBracketAccess()
    {
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[\"Price\"] > 5)");

        Assert.Contains("GetHeadersDelegate", result.SourceCode);
    }

    [Fact]
    public void Emit_NoHeaders_WhenNoStringBracketAccess()
    {
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.DoesNotContain("GetHeadersDelegate", result.SourceCode);
    }

    [Fact]
    public void Emit_UniformPreamble_AllParamsGetWrapping()
    {
        // Both tbl and threshold get the same wrapping treatment
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("tbl__isRef", result.SourceCode);
        Assert.Contains("threshold__isRef", result.SourceCode);
        Assert.Contains("ExcelValue.Wrap(tbl__values", result.SourceCode);
        Assert.Contains("ExcelValue.Wrap(threshold__values", result.SourceCode);
    }

    [Fact]
    public void Emit_PerVariableHeaders_OnlyRelevantParamGetsHeaders()
    {
        // tbl uses r["Price"] so needs headers; threshold doesn't
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[\"Price\"] > threshold)");

        // tbl should have header extraction
        Assert.Contains("tbl__headers = tbl__values is object[,]", result.SourceCode);
        Assert.Contains("GetHeadersDelegate", result.SourceCode);
        // threshold should NOT have header extraction
        Assert.Contains("string[]? threshold__headers = null;", result.SourceCode);
    }

    #endregion

    #region Expression Body

    [Fact]
    public void Emit_ExpressionPassedThrough()
    {
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] > 5)");

        Assert.Contains("tbl.Rows.Where(r => (double)r[0] > 5)", result.SourceCode);
    }

    [Fact]
    public void Emit_NoIExcelRangeCast()
    {
        // IExcelRange cast has been removed — all params wrapped uniformly
        var result = EmitFromExpression("tbl.Rows.Count()");

        Assert.DoesNotContain("(IExcelRange)", result.SourceCode);
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

    #region Free Variables as Parameters

    [Fact]
    public void Emit_FreeVariable_BecomesParameter()
    {
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("object threshold__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_FreeVariable_WrappedUniformly()
    {
        var result = EmitFromExpression("tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("ExcelValue.Wrap(threshold__values", result.SourceCode);
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

        Assert.Equal($"{CodeEmitter.UdfPrefix}HELLOWORLD", name);
    }

    [Fact]
    public void GenerateMethodName_PrependUnderscoreForDigitStart()
    {
        var name = CodeEmitter.GenerateMethodName("x", "123abc");

        Assert.Equal($"{CodeEmitter.UdfPrefix}_123ABC", name);
    }

    [Fact]
    public void GenerateMethodName_EmptyPreferredName_UsesHash()
    {
        var name = CodeEmitter.GenerateMethodName("tbl.Rows.Count()", "   ");

        Assert.StartsWith(CodeEmitter.UdfPrefix, name);
        Assert.NotEqual(CodeEmitter.UdfPrefix, name);
    }

    #endregion

    #region Statement Blocks

    [Fact]
    public void Emit_StatementBlock_UsesLocalFunction()
    {
        var result = EmitFromExpression("{ var x = 1; return tbl.Sum() + x; }");

        Assert.Contains("object __userBlock()", result.SourceCode);
        Assert.Contains("var __result = __userBlock();", result.SourceCode);
    }

    [Fact]
    public void Emit_StatementBlock_ContainsUserBlock()
    {
        var result = EmitFromExpression("{ var x = 1; return tbl.Sum() + x; }");

        Assert.Contains("return tbl.Sum() + x;", result.SourceCode);
    }

    [Fact]
    public void Emit_StatementBlock_StillWrapsParameters()
    {
        var result = EmitFromExpression("{ return tbl.Sum(); }");

        Assert.Contains("tbl__raw", result.SourceCode);
        Assert.Contains("ExcelValue.Wrap(", result.SourceCode);
    }

    [Fact]
    public void Emit_StatementBlock_StillConvertsResult()
    {
        var result = EmitFromExpression("{ return tbl.Sum(); }");

        Assert.Contains("ToResultDelegate", result.SourceCode);
    }

    #endregion
}
