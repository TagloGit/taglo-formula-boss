using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaRewriterTests
{
    #region Basic Rewriting

    [Fact]
    public void Rewrite_PreservesNonBacktickBindings()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        Assert.Contains("x, 1,", result);
        Assert.Contains("y, 2,", result);
        Assert.Contains("x + y)", result);
    }

    [Fact]
    public void Rewrite_InsertsSrcVariable_ForSingleBacktickBinding()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", new[] { "data" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("_src_filtered", result);
        Assert.Contains("\"data.where(v => v > 0)\"", result);
        Assert.Contains("filtered, FILTERED(data)", result);
    }

    [Fact]
    public void Rewrite_HandlesMultipleBacktickBindings()
    {
        var formula =
            "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, doubled, `filtered.select(v => v * 2)`, SUM(doubled))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", new[] { "data" }),
            ["doubled"] = new("doubled", "filtered.select(v => v * 2)", "DOUBLED", new[] { "filtered" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("_src_filtered", result);
        Assert.Contains("_src_doubled", result);
        Assert.Contains("FILTERED(data)", result);
        Assert.Contains("DOUBLED(filtered)", result);
    }

    [Fact]
    public void Rewrite_PreservesResultExpression()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", new[] { "data" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.EndsWith("SUM(filtered))", result.TrimEnd());
    }

    [Fact]
    public void Rewrite_HandlesBacktickResultExpression()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, `filtered.max()`)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", new[] { "data" })
        };

        var processedResults = new[] { new ProcessedBinding("_result", "filtered.max()", "_RESULT", new[] { "filtered" }) };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings, processedResults);

        Assert.Contains("_src__result", result);
        Assert.Contains("\"filtered.max()\"", result);
        Assert.Contains("_result, _RESULT(filtered)", result);
        Assert.EndsWith("_result)", result.TrimEnd());
    }

    [Fact]
    public void Rewrite_HandlesMultipleBacktickResultExpressions()
    {
        var formula = "=LET(a, A1:A3, b, B1:B3, `a.Sum()` + `b.Sum()`)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedResults = new[]
        {
            new ProcessedBinding("_result_1", "a.Sum()", "_RESULT_1", new[] { "a" }),
            new ProcessedBinding("_result_2", "b.Sum()", "_RESULT_2", new[] { "b" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>(),
            processedResults, "_result_1 + _result_2");

        Assert.Contains("_src__result_1", result);
        Assert.Contains("\"a.Sum()\"", result);
        Assert.Contains("_result_1, _RESULT_1(a)", result);
        Assert.Contains("_src__result_2", result);
        Assert.Contains("\"b.Sum()\"", result);
        Assert.Contains("_result_2, _RESULT_2(b)", result);
        Assert.EndsWith("_result_1 + _result_2)", result.TrimEnd());
    }

    #endregion

    #region Mixed Bindings

    [Fact]
    public void Rewrite_HandlesMixedBacktickAndNormalBindings()
    {
        var formula = "=LET(raw, A1:A100, threshold, 50, filtered, `raw.where(v => v > threshold)`, COUNT(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "raw.where(v => v > threshold)", "FILTERED", new[] { "raw", "threshold" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("raw, A1:A100", result);
        Assert.Contains("threshold, 50", result);
        Assert.Contains("_src_filtered", result);
        Assert.Contains("FILTERED(raw, threshold)", result);
    }

    #endregion

    #region String Escaping

    [Fact]
    public void Rewrite_EscapesQuotesInExpression()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v == \"test\")`, filtered)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v == \"test\")", "FILTERED", new[] { "data" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("\"\"test\"\"", result);
    }

    #endregion

    #region Full Formula Examples

    [Fact]
    public void Rewrite_FullExample_ColorFiltering()
    {
        var formula = @"=LET(data, A1:F20,
     coloredCells, `data.cells.where(c => c.color != -4142)`,
     result, `coloredCells.Select(c => c.Value * 2)`,
     SUM(result))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["coloredCells"] =
                new("coloredCells", "data.cells.where(c => c.color != -4142)", "COLOREDCELLS", new[] { "data" }),
            ["result"] = new("result", "coloredCells.Select(c => c.Value * 2)", "RESULT",
                new[] { "coloredCells" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.StartsWith("=LET(", result);
        Assert.EndsWith("SUM(result))", result);
        Assert.Contains("_src_coloredCells", result);
        Assert.Contains("_src_result", result);
        Assert.Contains("COLOREDCELLS(data)", result);
        Assert.Contains("RESULT(coloredCells)", result);
        Assert.Contains("data.cells.where(c => c.color != -4142)", result);
        Assert.Contains("coloredCells.Select(c => c.Value * 2)", result);
    }

    #endregion

    #region Formatting

    [Fact]
    public void Rewrite_FormatsWithLineBreaks()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        var lines = result.Split('\n');
        Assert.True(lines.Length >= 3, "Expected at least 3 lines (header, bindings, result)");
    }

    [Fact]
    public void Rewrite_IndentsBindings()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        Assert.Contains("    x, 1,", result);
        Assert.Contains("    y, 2,", result);
    }

    [Fact]
    public void Rewrite_RespectsCustomIndentSize()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>(),
            indentSize: 2);

        Assert.Contains("  x, 1,", result);
        Assert.Contains("  y, 2,", result);
        // Should NOT contain 4-space indent
        Assert.DoesNotContain("    x,", result);
    }

    [Fact]
    public void Rewrite_DelegatesToLetFormulaFormatter()
    {
        // Verify that the rewriter produces the same output as the formatter
        // would for the equivalent flat formula
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>(),
            indentSize: 2);

        var expected = LetFormulaFormatter.Format("=LET(x, 1, y, 2, x + y)", indentSize: 2);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Rewrite_WithProcessedBindings_FormatsViaFormatter()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", new[] { "data" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings, indentSize: 2);

        // Should use 2-space indent
        Assert.Contains("  _src_filtered,", result);
        Assert.Contains("  data, A1:A10,", result);
    }

    [Fact]
    public void Rewrite_AlwaysFormatsAtLeastDepth1_EvenWhenNestedLetDepthIsZero()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>(),
            nestedLetDepth: 0);

        // Rewriter constructs formulas from scratch — always formats for readability
        var lines = result.Split('\n');
        Assert.True(lines.Length >= 3, "Expected formatted output even with nestedLetDepth=0");
    }

    #endregion

    #region Flat Parameters in UDF Calls

    [Fact]
    public void Rewrite_FlatParameters_PassedToUdf()
    {
        var formula = "=LET(data, A1:A10, maxVal, 10, filtered, `data.Rows.Where(r => r[0] > maxVal)`, filtered)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.Rows.Where(r => r[0] > maxVal)",
                "FILTERED", new[] { "data", "maxVal" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("FILTERED(data, maxVal)", result);
    }

    [Fact]
    public void Rewrite_NoParameters_EmptyParentheses()
    {
        var formula = "=LET(x, `1 + 2`, x)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["x"] = new("x", "1 + 2", "LITERAL", Array.Empty<string>())
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        Assert.Contains("LITERAL()", result);
    }

    #endregion
}
