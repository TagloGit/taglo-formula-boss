using FormulaBoss.Compilation;
using FormulaBoss.Interception;
using FormulaBoss.Transpilation;
using FormulaBoss.UI;

using Xunit;

namespace FormulaBoss.Tests;

public class InterceptionTests
{
    /// <summary>
    ///     Mock compiler for testing the pipeline without actual Roslyn compilation.
    /// </summary>
    private sealed class MockDynamicCompiler : DynamicCompiler
    {
        public int CompileCount { get; private set; }

        public override List<string> CompileAndRegister(string source, bool isMacroType = false)
        {
            CompileCount++;
            // Return empty list (success) without actually compiling
            return [];
        }
    }

    #region BacktickExtractor.IsBacktickFormula Tests

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForNull() => Assert.False(BacktickExtractor.IsBacktickFormula(null));

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForEmptyString() =>
        Assert.False(BacktickExtractor.IsBacktickFormula(""));

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForFormulaWithoutBackticks() =>
        // Formula-like text without backticks should not match
        Assert.False(BacktickExtractor.IsBacktickFormula("=SUM(A1:A10)"));

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForPlainText() =>
        Assert.False(BacktickExtractor.IsBacktickFormula("hello world"));

    [Fact]
    public void IsBacktickFormula_ReturnsTrue_ForFormulaWithBackticks() =>
        // Excel stores '=SUM(`...`) as =SUM(`...`) - the apostrophe is hidden
        Assert.True(BacktickExtractor.IsBacktickFormula("=SUM(`data.Where(x => x > 0)`)"));

    [Fact]
    public void IsBacktickFormula_ReturnsFalse_ForBackticksWithoutEquals() =>
        Assert.False(BacktickExtractor.IsBacktickFormula("`data.Where(x => x > 0)`"));

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
        var result = BacktickExtractor.Extract("=SUM(`data.Where(x => x > 0)`)");

        Assert.Single(result);
        Assert.Equal("data.Where(x => x > 0)", result[0].Expression);
    }

    [Fact]
    public void Extract_ReturnsMultipleExpressions()
    {
        var result = BacktickExtractor.Extract("=SUM(`a.Sum()`) + SUM(`b.Sum()`)");

        Assert.Equal(2, result.Count);
        Assert.Equal("a.Sum()", result[0].Expression);
        Assert.Equal("b.Sum()", result[1].Expression);
    }

    [Fact]
    public void Extract_TracksCorrectPositions()
    {
        var input = "=`expr`";
        var result = BacktickExtractor.Extract(input);

        Assert.Single(result);
        Assert.Equal(1, result[0].StartIndex); // Position of opening backtick
        Assert.Equal(7, result[0].EndIndex); // Position after closing backtick
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
        var result = BacktickExtractor.Extract("=SUM(`data.Where(x => x > 0))");

        Assert.Empty(result);
    }

    #endregion

    #region BacktickExtractor.RewriteFormula Tests

    [Fact]
    public void RewriteFormula_ReplacesSingleExpression()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.Where(x => x > 0)"] = $"{CodeEmitter.UdfPrefix}abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula("=SUM(`data.Where(x => x > 0)`)", replacements);

        Assert.Equal($"=SUM({CodeEmitter.UdfPrefix}abc(data))", result);
    }

    [Fact]
    public void RewriteFormula_ReplacesMultipleExpressions()
    {
        var replacements = new Dictionary<string, string>
        {
            ["a.Sum()"] = $"{CodeEmitter.UdfPrefix}1(a)",
            ["b.Sum()"] = $"{CodeEmitter.UdfPrefix}2(b)"
        };

        var result = BacktickExtractor.RewriteFormula("=SUM(`a.Sum()`) + SUM(`b.Sum()`)", replacements);

        Assert.Equal($"=SUM({CodeEmitter.UdfPrefix}1(a)) + SUM({CodeEmitter.UdfPrefix}2(b))", result);
    }

    [Fact]
    public void RewriteFormula_PreservesFormulaStructure()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.cells.where(c=>c.color==6).select(c=>c.value)"] = $"{CodeEmitter.UdfPrefix}abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula(
            "=AVERAGE(`data.cells.where(c=>c.color==6).select(c=>c.value)`)",
            replacements);

        Assert.Equal($"=AVERAGE({CodeEmitter.UdfPrefix}abc(data))", result);
    }

    [Fact]
    public void RewriteFormula_HandlesStandaloneExpression()
    {
        var replacements = new Dictionary<string, string>
        {
            ["data.Select(x => x * 2)"] = $"{CodeEmitter.UdfPrefix}abc(data)"
        };

        var result = BacktickExtractor.RewriteFormula("=`data.Select(x => x * 2)`", replacements);

        Assert.Equal($"={CodeEmitter.UdfPrefix}abc(data)", result);
    }

    #endregion

    #region FormulaPipeline Tests

    [Fact]
    public void Pipeline_ParsesSimpleExpression()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("data.Select(x => x * 2)");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
        Assert.StartsWith($"{CodeEmitter.UdfPrefix}", result.UdfName);
    }

    [Fact]
    public void Pipeline_UsesPreferredName_FromContext()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var context = new ExpressionContext("filteredData");

        var result = pipeline.Process("data.Where(v => v > 0)", context);

        Assert.True(result.Success);
        Assert.Equal($"{CodeEmitter.UdfPrefix}FILTEREDDATA", result.UdfName);
    }

    [Fact]
    public void Pipeline_CachesWithContext_Separately()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        // Same expression with different preferred names should create different UDFs
        var result1 = pipeline.Process("data.Select(x => x * 2)", new ExpressionContext("firstUdf"));
        var result2 = pipeline.Process("data.Select(x => x * 2)", new ExpressionContext("secondUdf"));

        Assert.True(result1.Success);
        Assert.True(result2.Success);
        Assert.Equal($"{CodeEmitter.UdfPrefix}FIRSTUDF", result1.UdfName);
        Assert.Equal($"{CodeEmitter.UdfPrefix}SECONDUDF", result2.UdfName);
        Assert.Equal(2, compiler.CompileCount); // Should compile twice
    }

    [Fact]
    public void Pipeline_ReturnsError_ForEmptyExpression()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Pipeline_ExtractsFlatParameters()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("myRange.Cells.Where(c => c.Color == 6)");

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("myRange", result.Parameters);
    }

    [Fact]
    public void Pipeline_CachesResults()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result1 = pipeline.Process("data.Select(x => x * 2)");
        var result2 = pipeline.Process("data.Select(x => x * 2)");

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

        var result = pipeline.Process("data.Cells.Where(c => c.Color == 6).Select(c => c.Value)");

        Assert.True(result.Success);
        Assert.NotNull(result.UdfName);
    }

    [Fact]
    public void Pipeline_AppendsAllToHeaderVariableTableParameter()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("tbl.Rows.Where(r => (double)r[\"Amount\"] > 150).Count()");

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("tbl[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_DoesNotAppendAllToRangeRefHeaderVariable()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("A1:B10.Rows.Where(r => (double)r[\"Price\"] > 5).Count()");

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        // Range refs should stay as-is (user controls the range)
        Assert.Contains("A1:B10", result.Parameters);
        Assert.DoesNotContain("A1:B10[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_DoesNotAppendAllToNonHeaderVariable()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process("tbl.Rows.Where(r => (double)r[0] > threshold).Count()");

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        // threshold is not a header variable — no [#All]
        Assert.Contains("threshold", result.Parameters);
        Assert.DoesNotContain("threshold[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_MetadataDetectsTableParameter_WhenAstMisses()
    {
        // Scenario: bracket access outside a lambda (AST pattern matching misses this)
        // e.g. tblCastDist.Rows.First(r => (double)r["Distance"] < 100)["Castle"]
        // The outer ["Castle"] is on the result, not inside a lambda tracing back to the param.
        // But metadata knows tblCastDist is a table, so it should get [#All].
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var metadata = new WorkbookMetadata(
            new[] { "tblCastDist" },
            Array.Empty<string>(),
            new Dictionary<string, IReadOnlyList<string>>
            {
                ["tblCastDist"] = new[] { "Castle", "Distance" }
            });
        var context = new ExpressionContext(null, metadata);

        var result = pipeline.Process(
            "tblCastDist.Rows.Count()", context);

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("tblCastDist[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_MetadataDoesNotAffectNonTableParameters()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var metadata = new WorkbookMetadata(
            new[] { "tblSales" },
            Array.Empty<string>(),
            new Dictionary<string, IReadOnlyList<string>>());
        var context = new ExpressionContext(null, metadata);

        var result = pipeline.Process("threshold.ToString()", context);

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("threshold", result.Parameters);
        Assert.DoesNotContain("threshold[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_MetadataTableDetection_IsCaseInsensitive()
    {
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var metadata = new WorkbookMetadata(
            new[] { "tblCastles" },
            Array.Empty<string>(),
            new Dictionary<string, IReadOnlyList<string>>());
        var context = new ExpressionContext(null, metadata);

        // Parameter name uses different casing than metadata
        var result = pipeline.Process("tblcastles.Rows.Count()", context);

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("tblcastles[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_MetadataDoesNotOverrideRangeRefMapping()
    {
        // If a parameter is a range ref (e.g. A1:B10), it should stay as-is
        // even if its placeholder name somehow matches a table name
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);
        var metadata = new WorkbookMetadata(
            Array.Empty<string>(),
            Array.Empty<string>(),
            new Dictionary<string, IReadOnlyList<string>>());
        var context = new ExpressionContext(null, metadata);

        var result = pipeline.Process("A1:B10.Rows.Where(r => (double)r[\"Price\"] > 5).Count()", context);

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("A1:B10", result.Parameters);
        Assert.DoesNotContain("A1:B10[#All]", result.Parameters);
    }

    [Fact]
    public void Pipeline_NullMetadata_FallsBackToAstDetection()
    {
        // Without metadata, the existing AST-based detection should still work
        var compiler = new MockDynamicCompiler();
        var pipeline = new FormulaPipeline(compiler);

        var result = pipeline.Process(
            "tbl.Rows.Where(r => (double)r[\"Amount\"] > 150).Count()");

        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
        Assert.Contains("tbl[#All]", result.Parameters);
    }

    #endregion
}
