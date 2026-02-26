using FormulaBoss.Compilation;
using FormulaBoss.Interception;
using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Integration tests: expression → InputDetector → CodeEmitter → generated C# compiles.
///     Uses a mock DynamicCompiler that validates compilation without ExcelDNA registration.
/// </summary>
public class PipelineIntegrationTests
{
    // === InputDetector → CodeEmitter integration ===

    [Fact]
    public void SingleInput_Sugar_GeneratesCompilableCode()
    {
        var expr = "tblCountries.Rows.Where(r => r.Population > 1000)";
        var detection = InputDetector.Detect(expr);
        var result = CodeEmitter.Emit(detection, "test", expr);

        Assert.NotEmpty(result.SourceCode);
        Assert.Equal("TEST", result.MethodName);
        Assert.False(result.RequiresObjectModel);
    }

    [Fact]
    public void ExplicitLambda_MultiInput_GeneratesCompilableCode()
    {
        var expr = "(tbl, maxPop) => tbl.Rows.Where(r => r.Population < maxPop)";
        var detection = InputDetector.Detect(expr);
        var result = CodeEmitter.Emit(detection, "multiTest", expr);

        Assert.Equal(2, detection.Inputs.Count);
        Assert.Contains("object tbl__raw", result.SourceCode);
        Assert.Contains("object maxPop__raw", result.SourceCode);
    }

    [Fact]
    public void StatementBlock_GeneratesCompilableCode()
    {
        var expr = "(tbl) => { return tbl.Rows.Count(); }";
        var detection = InputDetector.Detect(expr);
        var result = CodeEmitter.Emit(detection, "stmtTest", expr);

        Assert.True(detection.IsStatementBody);
        Assert.Contains("__exec()", result.SourceCode);
    }

    [Fact]
    public void CellUsage_SetsRequiresObjectModel()
    {
        var expr = "tbl.Rows.Where(r => r.Price.Cell.Color == 6)";
        var detection = InputDetector.Detect(expr);
        var result = CodeEmitter.Emit(detection, "cellTest", expr);

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void FreeVariable_WithLetContext_BecomesParam()
    {
        var expr = "tbl.Rows.Where(r => r.Population < maxPop)";
        var knownVars = new HashSet<string> { "maxPop" };
        var detection = InputDetector.Detect(expr, knownVars);
        var result = CodeEmitter.Emit(detection, "letTest", expr);

        Assert.Contains("maxPop", detection.FreeVariables);
        Assert.Contains("object maxPop__raw", result.SourceCode);
        Assert.NotNull(result.UsedColumnBindings);
        Assert.Contains("maxPop", result.UsedColumnBindings);
    }

    [Fact]
    public void NestedLambda_FreevarsCorrect()
    {
        var expr = "tbl.Rows.Where(r => pConts.Any(c => c == r.Continent))";
        var knownVars = new HashSet<string> { "pConts" };
        var detection = InputDetector.Detect(expr, knownVars);

        Assert.Contains("pConts", detection.FreeVariables);
        Assert.DoesNotContain("r", detection.FreeVariables);
        Assert.DoesNotContain("c", detection.FreeVariables);
    }

    // === FormulaPipeline integration (with mock compiler) ===

    [Fact]
    public void Pipeline_Process_ReturnsSuccess_WithMockCompiler()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("tblCountries.Rows.Where(r => r.Population > 1000)");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
        Assert.NotNull(result.InputParameter);
        Assert.Equal("tblCountries", result.InputParameter);
    }

    [Fact]
    public void Pipeline_CachesResults()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var expr = "tblCountries.Rows.Count()";
        var result1 = pipeline.Process(expr);
        var result2 = pipeline.Process(expr);

        Assert.True(result1.Success);
        Assert.True(result2.Success);
        Assert.Equal(result1.UdfName, result2.UdfName);
        Assert.Equal(1, compiler.CompileCount);
    }

    [Fact]
    public void Pipeline_WithPreferredName()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var context = new ExpressionContext("myResult");
        var result = pipeline.Process("tbl.Rows.Count()", context);

        Assert.True(result.Success);
        Assert.Equal("MYRESULT", result.UdfName);
    }

    [Fact]
    public void Pipeline_ExplicitLambda()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("(tbl, maxPop) => tbl.Rows.Where(r => r.Population < maxPop)");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
    }

    [Fact]
    public void Pipeline_CompileError_ReturnsFailure()
    {
        var compiler = new MockDynamicCompiler { ShouldFail = true };
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("tbl.Rows.Count()");

        Assert.False(result.Success);
        Assert.Contains("Compile error", result.ErrorMessage);
    }

    private class MockDynamicCompiler : DynamicCompiler
    {
        public int CompileCount { get; private set; }
        public bool ShouldFail { get; set; }
        public string? LastSource { get; private set; }

        public override List<string> CompileAndRegister(string source, bool isMacroType = false)
        {
            CompileCount++;
            LastSource = source;

            if (ShouldFail)
            {
                return new List<string> { "Mock compilation failure" };
            }

            return new List<string>();
        }
    }
}
