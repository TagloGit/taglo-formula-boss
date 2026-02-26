using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class CodeEmitterTests
{
    [Fact]
    public void Emit_SugarSyntax_GeneratesWrapAndResult()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "tblCountries" },
            Body: "tblCountries.Rows.Where(r => r.Population > 1000)",
            IsExplicitLambda: false,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "myUdf", "tblCountries.Rows.Where(r => r.Population > 1000)");

        Assert.Contains("ExcelValue.Wrap(", result.SourceCode);
        Assert.Contains("ToResult(", result.SourceCode);
        Assert.Contains("public static object?[,] MYUDF(", result.SourceCode);
        Assert.Contains("object tblCountries__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_MultiInput_GeneratesMultipleParams()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "tbl", "maxPop" },
            Body: "tbl.Rows.Where(r => r.Population < maxPop)",
            IsExplicitLambda: true,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "myUdf", "original");

        Assert.Contains("object tbl__raw", result.SourceCode);
        Assert.Contains("object maxPop__raw", result.SourceCode);
    }

    [Fact]
    public void Emit_FreeVariables_BecomeAdditionalParams()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "tbl" },
            Body: "tbl.Rows.Where(r => r.X > threshold)",
            IsExplicitLambda: false,
            RequiresObjectModel: false,
            FreeVariables: new[] { "threshold" },
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "myUdf", "original");

        Assert.Contains("object threshold__raw", result.SourceCode);
        Assert.NotNull(result.UsedColumnBindings);
        Assert.Contains("threshold", result.UsedColumnBindings);
    }

    [Fact]
    public void Emit_RequiresObjectModel_FlowsThrough()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "tbl" },
            Body: "tbl.Rows.Where(r => r.X.Cell.Color == 6)",
            IsExplicitLambda: false,
            RequiresObjectModel: true,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "myUdf", "original");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Emit_StatementBody_UsesLocalFunction()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "tbl" },
            Body: "{ return tbl.Rows.Count(); }",
            IsExplicitLambda: true,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: true);

        var result = CodeEmitter.Emit(detection, "myUdf", "original");

        Assert.Contains("__exec()", result.SourceCode);
        Assert.Contains("return tbl.Rows.Count();", result.SourceCode);
    }

    [Fact]
    public void Emit_MethodName_Sanitized()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "x" },
            Body: "x.Rows.Count()",
            IsExplicitLambda: false,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "my-udf!", "original");

        Assert.Equal("MYUDF", result.MethodName);
    }

    [Fact]
    public void Emit_ReservedExcelName_Prefixed()
    {
        Assert.Equal("_SUM", CodeEmitter.SanitizeUdfName("SUM"));
        Assert.Equal("_FILTER", CodeEmitter.SanitizeUdfName("FILTER"));
    }

    [Fact]
    public void Emit_GeneratesUsings()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "x" },
            Body: "x.Rows.Count()",
            IsExplicitLambda: false,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "test", "original");

        Assert.Contains("using System;", result.SourceCode);
        Assert.Contains("using System.Linq;", result.SourceCode);
        Assert.Contains("using FormulaBoss.Runtime;", result.SourceCode);
    }

    [Fact]
    public void Emit_ResolvesExcelReference()
    {
        var detection = new InputDetectionResult(
            Inputs: new[] { "data" },
            Body: "data.Rows.Count()",
            IsExplicitLambda: false,
            RequiresObjectModel: false,
            FreeVariables: Array.Empty<string>(),
            IsStatementBody: false);

        var result = CodeEmitter.Emit(detection, "test", "original");

        Assert.Contains("GetValuesFromReference", result.SourceCode);
        Assert.Contains("ExcelReference", result.SourceCode);
    }
}
