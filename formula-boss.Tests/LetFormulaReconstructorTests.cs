using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaReconstructorTests
{
    #region IsProcessedFormulaBossLet

    [Fact]
    public void IsProcessedFormulaBossLet_ReturnsFalse_ForNullOrEmpty()
    {
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet(null));
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet(""));
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet("   "));
    }

    [Fact]
    public void IsProcessedFormulaBossLet_ReturnsFalse_ForNonLetFormula()
    {
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet("=SUM(A1:A10)"));
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet("=A1+B1"));
    }

    [Fact]
    public void IsProcessedFormulaBossLet_ReturnsFalse_ForNormalLet()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";
        Assert.False(LetFormulaReconstructor.IsProcessedFormulaBossLet(formula));
    }

    [Fact]
    public void IsProcessedFormulaBossLet_ReturnsTrue_ForProcessedLet()
    {
        var formula = @"=LET(data, A1:F20,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, FILTERED(data),
            SUM(filtered))";
        Assert.True(LetFormulaReconstructor.IsProcessedFormulaBossLet(formula));
    }

    [Fact]
    public void IsProcessedFormulaBossLet_ReturnsTrue_ForMultipleSrcBindings()
    {
        var formula = @"=LET(
            data, A1:F20,
            _src_coloredCells, ""data.cells.where(c => c.color != -4142)"",
            coloredCells, COLOREDCELLS(data),
            _src_result, ""coloredCells.Select(c => c.Value * 2)"",
            result, RESULT(coloredCells),
            SUM(result))";
        Assert.True(LetFormulaReconstructor.IsProcessedFormulaBossLet(formula));
    }

    #endregion

    #region TryReconstruct Basic

    [Fact]
    public void TryReconstruct_ReturnsFalse_ForNonFormulaBossFormula()
    {
        var formula = "=LET(x, 1, x)";
        var result = LetFormulaReconstructor.TryReconstruct(formula, out var editable);
        Assert.False(result);
        Assert.Null(editable);
    }

    [Fact]
    public void TryReconstruct_ReturnsFalse_ForNonLetFormula()
    {
        var formula = "=SUM(A1:A10)";
        var result = LetFormulaReconstructor.TryReconstruct(formula, out var editable);
        Assert.False(result);
        Assert.Null(editable);
    }

    [Fact]
    public void TryReconstruct_AddsQuotePrefix()
    {
        var processed = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, FILTERED(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.StartsWith("'", editable);
    }

    #endregion

    #region Single Binding Reconstruction

    [Fact]
    public void TryReconstruct_ReconstructsSingleBacktickBinding()
    {
        var processed = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, FILTERED(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.StartsWith("'=LET(", editable);
        Assert.Contains("`data.where(v => v > 0)`", editable);
        Assert.DoesNotContain("_src_", editable);
        Assert.DoesNotContain("FILTERED(data)", editable);
    }

    [Fact]
    public void TryReconstruct_PreservesNormalBindings()
    {
        var processed = @"=LET(raw, A1:A100, threshold, 50,
            _src_filtered, ""raw.where(v => v > threshold)"",
            filtered, FILTERED(raw),
            COUNT(filtered))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("raw, A1:A100", editable);
        Assert.Contains("threshold, 50", editable);
    }

    [Fact]
    public void TryReconstruct_PreservesResultExpression()
    {
        var processed = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, FILTERED(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.EndsWith("SUM(filtered))", editable);
    }

    #endregion

    #region Multiple Bindings Reconstruction

    [Fact]
    public void TryReconstruct_ReconstructsMultipleBacktickBindings()
    {
        var processed = @"=LET(data, A1:F20,
            _src_coloredCells, ""data.cells.where(c => c.color != -4142)"",
            coloredCells, COLOREDCELLS(data),
            _src_result, ""coloredCells.Select(c => c.Value * 2)"",
            result, RESULT(coloredCells),
            SUM(result))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("`data.cells.where(c => c.color != -4142)`", editable);
        Assert.Contains("`coloredCells.Select(c => c.Value * 2)`", editable);
        Assert.DoesNotContain("COLOREDCELLS(data)", editable);
        Assert.DoesNotContain("RESULT(coloredCells)", editable);
    }

    [Fact]
    public void TryReconstruct_HandlesMixedBacktickAndNormalBindings()
    {
        var processed = @"=LET(raw, A1:A100, threshold, 50,
            _src_filtered, ""raw.where(v => v > threshold)"",
            filtered, FILTERED(raw),
            multiplier, 2,
            COUNT(filtered))";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("raw, A1:A100", editable);
        Assert.Contains("threshold, 50", editable);
        Assert.Contains("`raw.where(v => v > threshold)`", editable);
        Assert.Contains("multiplier, 2", editable);
    }

    #endregion

    #region String Unescaping

    [Fact]
    public void TryReconstruct_UnescapesQuotesInDslExpression()
    {
        var processed = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v == """"test"""")"",
            filtered, FILTERED(data),
            filtered)";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        // Should have single quotes after unescaping
        Assert.Contains(@"v == ""test""", editable);
    }

    #endregion

    #region Result Expression Handling

    [Fact]
    public void TryReconstruct_PreservesVariableReferenceAsResultExpression()
    {
        // Result expression is just "maxYellow" (a variable reference), NOT a backtick expression
        // This should NOT be replaced with a backtick, even though maxYellow has a _src_ entry
        var processed = @"=LET(
            data, B3:F11,
            _src_yellowCells, ""data.cells.where(c => c.color == 6)"",
            yellowCells, YELLOWCELLS(data),
            _src_maxYellow, ""yellowCells.max()"",
            maxYellow, MAXYELLOW(yellowCells),
            maxYellow)";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        // Result should be the variable name, not a backtick expression
        Assert.EndsWith("maxYellow)", editable);
        // Should NOT contain _result
        Assert.DoesNotContain("_result", editable);
    }

    [Fact]
    public void TryReconstruct_HandlesBacktickResultExpression()
    {
        // _result is treated as a normal variable name — same as if the user named it _result
        var processed = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, FILTERED(data),
            _src__result, ""filtered.max()"",
            _result, _RESULT(filtered),
            _result)";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("`data.where(v => v > 0)`", editable);
        Assert.Contains("_result, `filtered.max()`", editable);
        Assert.EndsWith("_result)", editable);
    }

    [Fact]
    public void TryReconstruct_HandlesMultipleBacktickResultExpressions()
    {
        // _result_N bindings are normal variables — result expression references them by name
        var processed = @"=LET(a, A1:A3, b, B1:B3,
            _src__result_1, ""a.Sum()"",
            _result_1, _RESULT_1(a),
            _src__result_2, ""b.Sum()"",
            _result_2, _RESULT_2(b),
            _result_1 + _result_2)";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("_result_1, `a.Sum()`", editable);
        Assert.Contains("_result_2, `b.Sum()`", editable);
        Assert.EndsWith("_result_1 + _result_2)", editable);
    }

    [Fact]
    public void TryReconstruct_AutoResultWithNoExplicitVariable()
    {
        // Regression: issue #202 — _result is just a normal binding, not expanded in result position
        var processed = @"=LET(
            steps, SEQUENCE(10),
            _src__result, ""steps.Select(s => s*3)"",
            _result, __FB__RESULT(steps),
            _result)";

        var result = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(result);
        Assert.NotNull(editable);
        Assert.Contains("steps, SEQUENCE(10)", editable);
        Assert.Contains("_result, `steps.Select(s => s*3)`", editable);
        Assert.EndsWith("_result)", editable);
        Assert.DoesNotContain("_src_", editable);
    }

    #endregion

    #region Header Binding Stripping

    [Fact]
    public void TryReconstruct_StripsHeaderBindings()
    {
        // Header bindings (_*_hdr) are injected machinery for dynamic column names
        // They should be stripped when reconstructing the editable formula
        var processed = @"=LET(
            tbl, tblSales,
            price, tblSales[Price],
            qty, tblSales[Qty],
            _price_hdr, INDEX(tblSales[[#Headers],[Price]],1),
            _qty_hdr, INDEX(tblSales[[#Headers],[Qty]],1),
            _src_result, ""tbl.Rows.Select(r => r.Price * r.Qty).Sum()"",
            result, RESULT(tbl, _price_hdr, _qty_hdr),
            result)";

        var success = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(success);
        Assert.NotNull(editable);
        // Header bindings should be stripped
        Assert.DoesNotContain("_price_hdr", editable);
        Assert.DoesNotContain("_qty_hdr", editable);
        Assert.DoesNotContain("[[#Headers]", editable);
        Assert.DoesNotContain("INDEX(", editable);
        // Original LET bindings should be preserved
        Assert.Contains("tbl, tblSales", editable);
        Assert.Contains("price, tblSales[Price]", editable);
        Assert.Contains("qty, tblSales[Qty]", editable);
        // Backtick expression should be reconstructed
        Assert.Contains("`tbl.Rows.Select(r => r.Price * r.Qty).Sum()`", editable);
    }

    [Fact]
    public void TryReconstruct_StripsMultipleHeaderBindings()
    {
        var processed = @"=LET(
            tbl, tblProducts,
            name, tblProducts[Name],
            price, tblProducts[Price],
            qty, tblProducts[Quantity],
            _name_hdr, INDEX(tblProducts[[#Headers],[Name]],1),
            _price_hdr, INDEX(tblProducts[[#Headers],[Price]],1),
            _qty_hdr, INDEX(tblProducts[[#Headers],[Quantity]],1),
            _src_total, ""tbl.Rows.Select(r => r.Price * r.Qty).Sum()"",
            total, TOTAL(tbl, _name_hdr, _price_hdr, _qty_hdr),
            total)";

        var success = LetFormulaReconstructor.TryReconstruct(processed, out var editable);

        Assert.True(success);
        Assert.NotNull(editable);
        // All header bindings should be stripped
        Assert.DoesNotContain("_name_hdr", editable);
        Assert.DoesNotContain("_price_hdr", editable);
        Assert.DoesNotContain("_qty_hdr", editable);
    }

    #endregion

    #region GetDebugCallSites

    [Fact]
    public void GetDebugCallSites_ReturnsEmpty_ForNullOrEmpty()
    {
        Assert.Empty(LetFormulaReconstructor.GetDebugCallSites(null));
        Assert.Empty(LetFormulaReconstructor.GetDebugCallSites(""));
    }

    [Fact]
    public void GetDebugCallSites_ReturnsEmpty_ForNormalFormula()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED(data),
            SUM(filtered))";
        Assert.Empty(LetFormulaReconstructor.GetDebugCallSites(formula));
    }

    [Fact]
    public void GetDebugCallSites_DetectsDebugCallSite()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED_DEBUG(data),
            SUM(filtered))";
        var names = LetFormulaReconstructor.GetDebugCallSites(formula);
        Assert.Single(names);
        Assert.Equal("FILTERED", names[0]);
    }

    [Fact]
    public void GetDebugCallSites_DetectsMultipleDebugCallSites()
    {
        var formula = @"=LET(data, A1:F20,
            _src_colored, ""data.cells.where(c => c.color != 0)"",
            colored, __FB_COLORED_DEBUG(data),
            _src_result, ""colored.max()"",
            result, __FB_RESULT_DEBUG(colored),
            result)";
        var names = LetFormulaReconstructor.GetDebugCallSites(formula);
        Assert.Equal(2, names.Count);
        Assert.Contains("COLORED", names);
        Assert.Contains("RESULT", names);
    }

    [Fact]
    public void GetDebugCallSites_IgnoresDebugInsideStringLiterals()
    {
        // The _src_ binding value is a string literal — should not be scanned
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""__FB_FILTERED_DEBUG(data)"",
            filtered, __FB_FILTERED(data),
            SUM(filtered))";
        Assert.Empty(LetFormulaReconstructor.GetDebugCallSites(formula));
    }

    #endregion

    #region RewriteCallSitesToDebug

    [Fact]
    public void RewriteCallSitesToDebug_RewritesSingleCallSite()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.RewriteCallSitesToDebug(formula, new[] { "FILTERED" });

        Assert.Contains("__FB_FILTERED_DEBUG(data)", result);
        Assert.DoesNotContain("__FB_FILTERED(data)", result);
    }

    [Fact]
    public void RewriteCallSitesToDebug_PreservesSourceLiterals()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.RewriteCallSitesToDebug(formula, new[] { "FILTERED" });

        // _src_ binding value should be untouched
        Assert.Contains(@"_src_filtered, ""data.where(v => v > 0)""", result);
    }

    [Fact]
    public void RewriteCallSitesToDebug_RewritesMultipleCallSites()
    {
        var formula = @"=LET(data, A1:F20,
            _src_colored, ""data.cells.where(c => c.color != 0)"",
            colored, __FB_COLORED(data),
            _src_result, ""colored.max()"",
            result, __FB_RESULT(colored),
            result)";

        var result = LetFormulaReconstructor.RewriteCallSitesToDebug(
            formula, new[] { "COLORED", "RESULT" });

        Assert.Contains("__FB_COLORED_DEBUG(data)", result);
        Assert.Contains("__FB_RESULT_DEBUG(colored)", result);
    }

    [Fact]
    public void RewriteCallSitesToDebug_DoesNotTouchUnspecifiedNames()
    {
        var formula = @"=LET(
            colored, __FB_COLORED(data),
            result, __FB_RESULT(colored),
            result)";

        // Only rewrite COLORED, not RESULT
        var result = LetFormulaReconstructor.RewriteCallSitesToDebug(formula, new[] { "COLORED" });

        Assert.Contains("__FB_COLORED_DEBUG(data)", result);
        Assert.Contains("__FB_RESULT(colored)", result);
        Assert.DoesNotContain("__FB_RESULT_DEBUG", result);
    }

    #endregion

    #region RewriteCallSitesToNormal

    [Fact]
    public void RewriteCallSitesToNormal_RewritesDebugCallSite()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED_DEBUG(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.RewriteCallSitesToNormal(formula, new[] { "FILTERED" });

        Assert.Contains("__FB_FILTERED(data)", result);
        Assert.DoesNotContain("_DEBUG", result);
    }

    [Fact]
    public void RewriteCallSitesToNormal_PreservesSourceLiterals()
    {
        var formula = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED_DEBUG(data),
            SUM(filtered))";

        var result = LetFormulaReconstructor.RewriteCallSitesToNormal(formula, new[] { "FILTERED" });

        Assert.Contains(@"_src_filtered, ""data.where(v => v > 0)""", result);
    }

    #endregion

    #region Round-trip Rewriting

    [Fact]
    public void RoundTrip_DebugThenNormal_RestoresOriginalFormula()
    {
        var original = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v > 0)"",
            filtered, __FB_FILTERED(data),
            SUM(filtered))";

        var debugged = LetFormulaReconstructor.RewriteCallSitesToDebug(original, new[] { "FILTERED" });
        var restored = LetFormulaReconstructor.RewriteCallSitesToNormal(debugged, new[] { "FILTERED" });

        Assert.Equal(original, restored);
    }

    [Fact]
    public void RoundTrip_MultipleNames_RestoresOriginalFormula()
    {
        var original = @"=LET(data, A1:F20,
            _src_colored, ""data.cells.where(c => c.color != 0)"",
            colored, __FB_COLORED(data),
            _src_result, ""colored.max()"",
            result, __FB_RESULT(colored),
            result)";

        var names = new[] { "COLORED", "RESULT" };
        var debugged = LetFormulaReconstructor.RewriteCallSitesToDebug(original, names);
        var restored = LetFormulaReconstructor.RewriteCallSitesToNormal(debugged, names);

        Assert.Equal(original, restored);
    }

    [Fact]
    public void RoundTrip_PreservesEscapedQuotesInSourceLiteral()
    {
        var original = @"=LET(data, A1:A10,
            _src_filtered, ""data.where(v => v == """"test"""")"",
            filtered, __FB_FILTERED(data),
            filtered)";

        var debugged = LetFormulaReconstructor.RewriteCallSitesToDebug(original, new[] { "FILTERED" });

        // Escaped quotes inside _src_ binding should be unchanged
        Assert.Contains(@"""""test""""", debugged);

        var restored = LetFormulaReconstructor.RewriteCallSitesToNormal(debugged, new[] { "FILTERED" });
        Assert.Equal(original, restored);
    }

    #endregion
}
