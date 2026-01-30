using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class InterceptionTests
{
    #region BacktickExtractor.IsBacktickFormula Tests

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForNull()
    {
        Assert.False(BacktickExtractor.IsBacktickFormula(null));
    }

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForEmptyString()
    {
        Assert.False(BacktickExtractor.IsBacktickFormula(""));
    }

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForFormulaWithoutBackticks()
    {
        // Formula-like text without backticks should not match
        Assert.False(BacktickExtractor.IsBacktickFormula("=SUM(A1:A10)"));
    }

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForPlainText()
    {
        Assert.False(BacktickExtractor.IsBacktickFormula("hello world"));
    }

    [Fact]
    public void IsBacktickFormula_ReturnsTrue_ForFormulaWithBackticks()
    {
        // Excel stores '=SUM(`...`) as =SUM(`...`) - the apostrophe is hidden
        Assert.True(BacktickExtractor.IsBacktickFormula("=SUM(`data.values`)"));
    }

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForBackticksWithoutEquals()
    {
        Assert.False(BacktickExtractor.IsBacktickFormula("`data.values`"));
    }

    #endregion

    #region BacktickExtractor.Extract Tests

    [Fact]
    public void Extract_ReturnsEmpty_ForNoBackticks()
    {
        var result = BacktickExtractor.Extract("=SUM(A1:A10)");

        Assert.Empty(result);
    }

    [Fact]
    public void Extract_ReturnsSingleExpression()
    {
        var result = BacktickExtractor.Extract("=SUM(`data.values`)");

        Assert.Single(result);
        Assert.Equal("data.values", result[0].Expression);
    }

    [Fact]
    public void Extract_ReturnsMultipleExpressions()
    {
        var result = BacktickExtractor.Extract("=SUM(`a.values`) + SUM(`b.values`)");

        Assert.Equal(2, result.Count);
        Assert.Equal("a.values", result[0].Expression);
        Assert.Equal("b.values", result[1].Expression);
    }

    [Fact]
    public void Extract_TracksCorrectPositions()
    {
        var input = "=`expr`";
        var result = BacktickExtractor.Extract(input);

        Assert.Single(result);
        Assert.Equal(1, result[0].StartIndex); // Position of opening backtick
        Assert.Equal(7, result[0].EndIndex);   // Position after closing backtick
    }

    [Fact]
    public void Extract_HandlesComplexExpression()
    {
        var result = BacktickExtractor.Extract("=SUM(`data.cells.where(c=>c.color==6).select(c=>c.value)`)");

        Assert.Single(result);
        Assert.Equal("data.cells.where(c=>c.color==6).select(c=>c.value)", result[0].Expression);
    }

    [Fact]
    public void Extract_IgnoresUnclosedBacktick()
    {
        var result = BacktickExtractor.Extract("=SUM(`data.values)");

        Assert.Empty(result);
    }

    #endregion

    #region BacktickExtractor.RewriteFormula Tests

    [Fact]
    public void RewriteFormula_ReplacesSingleExpression()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.values"] = "__udf_abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula("=SUM(`data.values`)", replacements);

        Assert.Equal("=SUM(__udf_abc(data))", result);
    }

    [Fact]
    public void RewriteFormula_ReplacesMultipleExpressions()
    {
        var replacements = new Dictionary<string, string>
        {
            ["a.values"] = "__udf_1(a)",
            ["b.values"] = "__udf_2(b)"
        };

        var result = BacktickExtractor.RewriteFormula("=SUM(`a.values`) + SUM(`b.values`)", replacements);

        Assert.Equal("=SUM(__udf_1(a)) + SUM(__udf_2(b))", result);
    }

    [Fact]
    public void RewriteFormula_PreservesFormulaStructure()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.cells.where(c=>c.color==6).select(c=>c.value)"] = "__udf_abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula(
            "=AVERAGE(`data.cells.where(c=>c.color==6).select(c=>c.value)`)",
            replacements);

        Assert.Equal("=AVERAGE(__udf_abc(data))", result);
    }

    [Fact]
    public void RewriteFormula_HandlesStandaloneExpression()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.values.toArray()"] = "__udf_abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula("=`data.values.toArray()`", replacements);

        Assert.Equal("=__udf_abc(data)", result);
    }

    #endregion

    #region FormulaPipeline Tests

    [Fact]
    public void Pipeline_ParsesSimpleExpression()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("data.values.toArray()");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
        Assert.StartsWith("__udf_", result.UdfName);
    }

    [Fact]
    public void Pipeline_UsesPreferredName_FromContext()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var context = new ExpressionContext("filteredData");

        var result = pipeline.Process("data.values.where(v => v > 0)", context);

        Assert.True(result.Success);
        Assert.Equal("FILTEREDDATA", result.UdfName);
    }

    [Fact]
    public void Pipeline_CachesWithContext_Separately()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        // Same expression with different preferred names should create different UDFs
        var result1 = pipeline.Process("data.values.toArray()", new ExpressionContext("firstUdf"));
        var result2 = pipeline.Process("data.values.toArray()", new ExpressionContext("secondUdf"));

        Assert.True(result1.Success);
        Assert.True(result2.Success);
        Assert.Equal("FIRSTUDF", result1.UdfName);
        Assert.Equal("SECONDUDF", result2.UdfName);
        Assert.Equal(2, compiler.CompileCount); // Should compile twice
    }

    [Fact]
    public void Pipeline_ReturnsError_ForInvalidSyntax()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("data.");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("error", result.ErrorMessage.ToLowerInvariant());
    }

    [Fact]
    public void Pipeline_ExtractsInputParameter()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("myRange.cells.toArray()");

        Assert.True(result.Success);
        Assert.Equal("myRange", result.InputParameter);
    }

    [Fact]
    public void Pipeline_CachesResults()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result1 = pipeline.Process("data.values.toArray()");
        var result2 = pipeline.Process("data.values.toArray()");

        Assert.True(result1.Success);
        Assert.True(result2.Success);
        Assert.Equal(result1.UdfName, result2.UdfName);
        Assert.Equal(1, compiler.CompileCount); // Should only compile once
    }

    [Fact]
    public void Pipeline_HandlesComplexExpression()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("data.cells.where(c => c.color == 6).select(c => c.value).toArray()");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
    }

    #endregion

    /// <summary>
    /// Mock compiler for testing the pipeline without actual Roslyn compilation.
    /// </summary>
    private sealed class MockDynamicCompiler : Compilation.DynamicCompiler
    {
        public int CompileCount { get; private set; }

        public override List<string> CompileAndRegister(string source)
        {
            CompileCount++;
            // Return empty list (success) without actually compiling
            return [];
        }
    }
}
