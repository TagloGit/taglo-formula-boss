using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaRewriterTests
{
    #region Mixed Bindings

    [Fact]
    public void Rewrite_HandlesMixedBacktickAndNormalBindings()
    {
        var formula = "=LET(raw, A1:A100, threshold, 50, filtered, `raw.where(v => v > threshold)`, COUNT(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "raw.where(v => v > threshold)", "FILTERED", "raw")
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Normal bindings preserved
        Assert.Contains("raw, A1:A100", result);
        Assert.Contains("threshold, 50", result);

        // Backtick binding converted
        Assert.Contains("_src_filtered", result);
        Assert.Contains("FILTERED(raw)", result);
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
            ["filtered"] = new("filtered", "data.where(v => v == \"test\")", "FILTERED", "data")
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Quotes should be doubled for Excel string escaping
        Assert.Contains("\"\"test\"\"", result);
    }

    #endregion

    #region Full Formula Examples

    [Fact]
    public void Rewrite_FullExample_ColorFiltering()
    {
        var formula = @"=LET(data, A1:F20,
     coloredCells, `data.cells.where(c => c.color != -4142)`,
     result, `coloredCells.select(c => c.value * 2).toArray()`,
     SUM(result))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["coloredCells"] =
                new("coloredCells", "data.cells.where(c => c.color != -4142)", "COLOREDCELLS", "data"),
            ["result"] = new("result", "coloredCells.select(c => c.value * 2).toArray()", "RESULT", "coloredCells")
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Verify structure
        Assert.StartsWith("=LET(", result);
        Assert.EndsWith("SUM(result))", result);

        // Verify _src_ variables
        Assert.Contains("_src_coloredCells", result);
        Assert.Contains("_src_result", result);

        // Verify UDF calls
        Assert.Contains("COLOREDCELLS(data)", result);
        Assert.Contains("RESULT(coloredCells)", result);

        // Verify original expressions preserved in strings
        Assert.Contains("data.cells.where(c => c.color != -4142)", result);
        Assert.Contains("coloredCells.select(c => c.value * 2).toArray()", result);
    }

    #endregion

    #region Basic Rewriting

    [Fact]
    public void Rewrite_PreservesNonBacktickBindings()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        // Verify bindings are preserved (now formatted with line breaks)
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
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", "data")
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
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", "data"),
            ["doubled"] = new("doubled", "filtered.select(v => v * 2)", "DOUBLED", "filtered")
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
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", "data")
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Result expression should be the last line
        Assert.EndsWith("SUM(filtered))", result.TrimEnd());
    }

    #endregion

    #region Formatting

    [Fact]
    public void Rewrite_FormatsWithLineBreaks()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        // Should have line breaks between bindings
        var lines = result.Split('\n');
        Assert.True(lines.Length >= 3, "Expected at least 3 lines (header, bindings, result)");
    }

    [Fact]
    public void Rewrite_IndentsBindings()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        LetFormulaParser.TryParse(formula, out var structure);

        var result = LetFormulaRewriter.Rewrite(structure!, new Dictionary<string, ProcessedBinding>());

        // Bindings should be indented
        Assert.Contains("    x, 1,", result);
        Assert.Contains("    y, 2,", result);
    }

    #endregion
}
