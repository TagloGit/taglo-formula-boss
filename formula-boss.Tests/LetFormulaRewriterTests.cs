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

    [Fact]
    public void Rewrite_HandlesBacktickResultExpression()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, `filtered.max()`)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", "data")
        };

        var processedResult = new ProcessedBinding("_result", "filtered.max()", "_RESULT", "filtered");

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings, processedResult);

        // Should have _src__result documentation
        Assert.Contains("_src__result", result);
        Assert.Contains("\"filtered.max()\"", result);

        // Should have _result binding with UDF call
        Assert.Contains("_result, _RESULT(filtered)", result);

        // Final expression should reference _result
        Assert.EndsWith("_result)", result.TrimEnd());
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

    #region Column Parameter Injection

    [Fact]
    public void Rewrite_InjectsHeaderBindings_ForColumnParameters()
    {
        var formula = "=LET(tbl, tblSales, price, tblSales[Price], result, `tbl.reduce(0, (acc, r) => acc + r.price)`, result)";
        LetFormulaParser.TryParse(formula, out var structure);

        var columnParams = new List<ColumnParameter>
        {
            new("price", "tblSales", "Price")
        };

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["result"] = new("result", "tbl.reduce(0, (acc, r) => acc + r.price)", "RESULT", "tbl", columnParams)
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Should inject _price_hdr header binding before the UDF call
        // Note: INDEX(...,1) forces value evaluation instead of returning a reference
        Assert.Contains("_price_hdr, INDEX(tblSales[[#Headers],[Price]],1)", result);
        // UDF call should include the header argument
        Assert.Contains("RESULT(tbl, _price_hdr)", result);
    }

    [Fact]
    public void Rewrite_InjectsMultipleHeaderBindings()
    {
        var formula = "=LET(tbl, tblSales, price, tblSales[Price], qty, tblSales[Qty], result, `tbl.reduce(0, (acc, r) => acc + r.price * r.qty)`, result)";
        LetFormulaParser.TryParse(formula, out var structure);

        var columnParams = new List<ColumnParameter>
        {
            new("price", "tblSales", "Price"),
            new("qty", "tblSales", "Qty")
        };

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["result"] = new("result", "tbl.reduce(0, (acc, r) => acc + r.price * r.qty)", "RESULT", "tbl", columnParams)
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Should inject both header bindings with INDEX() wrapper
        Assert.Contains("_price_hdr, INDEX(tblSales[[#Headers],[Price]],1)", result);
        Assert.Contains("_qty_hdr, INDEX(tblSales[[#Headers],[Qty]],1)", result);
        // UDF call should include both header arguments
        Assert.Contains("RESULT(tbl, _price_hdr, _qty_hdr)", result);
    }

    [Fact]
    public void Rewrite_NoHeaderBindings_WhenNoColumnParameters()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.where(v => v > 0)", "FILTERED", "data")
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Should not have any _xxx_hdr bindings
        Assert.DoesNotContain("_hdr,", result);
        Assert.DoesNotContain("[[#Headers]", result);
        // UDF call should just have the input parameter
        Assert.Contains("FILTERED(data)", result);
    }

    [Fact]
    public void Rewrite_DeduplicatesHeaderBindings_AcrossMultipleProcessedBindings()
    {
        var formula = "=LET(tbl, tblSales, price, tblSales[Price], filtered, `tbl.rows.where(r => r.price > 10)`, summed, `filtered.reduce(0, (acc, r) => acc + r.price)`, summed)";
        LetFormulaParser.TryParse(formula, out var structure);

        var priceParam = new List<ColumnParameter> { new("price", "tblSales", "Price") };

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "tbl.rows.where(r => r.price > 10)", "FILTERED", "tbl", priceParam),
            ["summed"] = new("summed", "filtered.reduce(0, (acc, r) => acc + r.price)", "SUMMED", "filtered", priceParam)
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // Should only inject _price_hdr once (before first usage)
        var headerBindingCount = result.Split("_price_hdr, INDEX(tblSales[[#Headers],[Price]],1)").Length - 1;
        Assert.Equal(1, headerBindingCount);
    }

    #endregion

    #region Multi-Input Lambdas

    [Fact]
    public void Rewrite_MultiInput_PassesAdditionalInputsToUdf()
    {
        var formula = "=LET(data, A1:A10, maxVal, 10, filtered, `(data, maxVal) => data.Rows.Where(r => r[0] > maxVal)`, filtered)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "(data, maxVal) => data.Rows.Where(r => r[0] > maxVal)",
                "FILTERED", "data", null, new List<string> { "maxVal" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // UDF call should include both inputs
        Assert.Contains("FILTERED(data, maxVal)", result);
    }

    [Fact]
    public void Rewrite_FreeVariables_PassedAsUdfArguments()
    {
        var formula = "=LET(data, A1:A10, threshold, 50, filtered, `data.Rows.Where(r => r[0] > threshold)`, filtered)";
        LetFormulaParser.TryParse(formula, out var structure);

        var processedBindings = new Dictionary<string, ProcessedBinding>
        {
            ["filtered"] = new("filtered", "data.Rows.Where(r => r[0] > threshold)",
                "FILTERED", "data", null, null, new List<string> { "threshold" })
        };

        var result = LetFormulaRewriter.Rewrite(structure!, processedBindings);

        // UDF call should include the free variable
        Assert.Contains("FILTERED(data, threshold)", result);
    }

    #endregion
}
