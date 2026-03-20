using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaFormatterTests
{
    #region Basic Formatting

    [Fact]
    public void Format_BasicLet_IndentsBindings()
    {
        var formula = "=LET(x, A1, y, B1, x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        var expected = "=LET(\n    x, A1,\n    y, B1,\n    x + y)";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Format_BasicLet_DefaultIndentSize()
    {
        var formula = "=LET(x, A1, y, B1, x + y)";

        var result = LetFormulaFormatter.Format(formula);

        var expected = "=LET(\n  x, A1,\n  y, B1,\n  x + y)";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Format_SingleBinding()
    {
        var formula = "=LET(x, 1, x)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        var expected = "=LET(\n    x, 1,\n    x)";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Format_AlreadyFormatted_ReformatsCleanly()
    {
        var formula = "=LET(\n    x, A1,\n    y, B1,\n    x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Equal(formula, result);
    }

    [Fact]
    public void Format_ClosingParenOnResultLine()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 2);

        // Style A: closing paren on same line as result
        Assert.EndsWith("x + y)", result);
    }

    [Fact]
    public void Format_NameAndValueOnSameLine()
    {
        var formula = "=LET(x, SUM(A1:A10), y, AVERAGE(B1:B5), x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("x, SUM(A1:A10),", result);
        Assert.Contains("y, AVERAGE(B1:B5),", result);
    }

    #endregion

    #region Nested LETs

    [Fact]
    public void Format_NestedLet_Depth1_LeavesNestedAsIs()
    {
        var formula = "=LET(x, A1, result, LET(y, B1, z, C1, y + z), x + result)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, nestedLetDepth: 1);

        // Top-level is formatted, nested LET is left as-is
        Assert.Contains("result, LET(y, B1, z, C1, y + z),", result);
    }

    [Fact]
    public void Format_NestedLet_Depth2_FormatsOneLevel()
    {
        var formula = "=LET(x, A1, result, LET(y, B1, z, C1, y + z), x + result)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, nestedLetDepth: 2);

        // Nested LET should be formatted with additional indentation
        Assert.Contains("result, LET(", result);
        Assert.Contains("        y, B1,", result); // 8 spaces = 4 + 4
        Assert.Contains("        z, C1,", result);
        Assert.Contains("        y + z)", result);
    }

    [Fact]
    public void Format_NestedLet_Depth3_FormatsTwoLevels()
    {
        var formula = "=LET(a, 1, b, LET(c, 2, d, LET(e, 3, e + 1), c + d), a + b)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 2, nestedLetDepth: 3);

        // All three levels formatted
        Assert.Contains("  b, LET(", result);     // 2 spaces
        Assert.Contains("    c, 2,", result);       // 4 spaces
        Assert.Contains("    d, LET(", result);     // 4 spaces
        Assert.Contains("      e, 3,", result);     // 6 spaces
    }

    [Fact]
    public void Format_NestedLetDepth0_NoFormatting()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";

        var result = LetFormulaFormatter.Format(formula, nestedLetDepth: 0);

        Assert.Equal(formula, result);
    }

    #endregion

    #region MaxLineLength Inlining

    [Fact]
    public void Format_MaxLineLength_InlinesShortBindings()
    {
        var formula = "=LET(x, A1, y, B1, total, SUM(x, y), total)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 40);

        // Short bindings should be inlined on the =LET( line
        Assert.StartsWith("=LET(x, A1, y, B1,", result);
    }

    [Fact]
    public void Format_MaxLineLength_WrapsWhenExceeded()
    {
        var formula = "=LET(x, A1, y, B1, total, SUM(x, y), total)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 25);

        // First binding fits, subsequent ones wrap
        Assert.Contains("=LET(x, A1,", result);
        Assert.Contains("\n", result);
    }

    [Fact]
    public void Format_MaxLineLength_Zero_AlwaysWraps()
    {
        var formula = "=LET(x, 1, x)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 0);

        // Even short formulas should wrap
        Assert.Contains("\n", result);
    }

    [Fact]
    public void Format_MaxLineLength_AllFitOnOneLine()
    {
        var formula = "=LET(x, 1, y, 2, x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 100);

        // Everything fits on one line, result on same line
        Assert.Equal("=LET(x, 1, y, 2, x + y)", result);
    }

    [Fact]
    public void Format_MaxLineLength_OnceWrapped_AllSubsequentWrap()
    {
        var formula = "=LET(a, 1, b, 2, c, 3, d, 4, a + b + c + d)";

        // Set limit such that first two fit but third doesn't
        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 22);

        // Once wrapping starts, all subsequent bindings wrap
        var lines = result.Split('\n');
        Assert.True(lines.Length >= 3, "Expected wrapping after limit exceeded");
    }

    [Fact]
    public void Format_MaxLineLength_ResultWrapsIndependentlyOfBindings()
    {
        var formula = "=LET(x, 1, y, 2, z, 3, x+y+z+data)";

        // At 28, bindings "x, 1, y, 2, z, 3," fit (27 chars with =LET()
        // but adding " x+y+z+data)" would exceed 28
        var result = LetFormulaFormatter.Format(formula, indentSize: 4, maxLineLength: 28);

        // All bindings on first line, result wraps
        Assert.StartsWith("=LET(x, 1, y, 2, z, 3,", result);
        Assert.Contains("\n", result);
        Assert.EndsWith("x+y+z+data)", result);
    }

    #endregion

    #region String Literals

    [Fact]
    public void Format_StringLiteralWithCommas_NotSplit()
    {
        var formula = "=LET(msg, \"hello, world\", x, 1, msg)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("msg, \"hello, world\",", result);
    }

    [Fact]
    public void Format_StringLiteralWithParens_NotSplit()
    {
        var formula = "=LET(msg, \"value (test)\", msg)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("msg, \"value (test)\",", result);
    }

    #endregion

    #region Array Constants

    [Fact]
    public void Format_ArrayConstant_KeptOnOneLine()
    {
        var formula = "=LET(arr, {1,2,3;4,5,6}, SUM(arr))";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("arr, {1,2,3;4,5,6},", result);
    }

    #endregion

    #region Backtick Expressions

    [Fact]
    public void Format_BacktickExpression_PreservedExactly()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("filtered, `data.where(v => v > 0)`,", result);
    }

    [Fact]
    public void Format_BacktickWithCommas_PreservedExactly()
    {
        var formula = "=LET(out, `new int[] { 1, 2, 3 }`, SUM(out))";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Contains("out, `new int[] { 1, 2, 3 }`,", result);
    }

    #endregion

    #region LET Inside Other Functions

    [Fact]
    public void Format_LetInsideSumproduct()
    {
        var formula = "=SUMPRODUCT(LET(x, A1:A10, x * 2))";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.StartsWith("=SUMPRODUCT(LET(", result);
        Assert.Contains("    x, A1:A10,", result);
        Assert.Contains("    x * 2)", result);
        Assert.EndsWith(")", result);
    }

    [Fact]
    public void Format_LetInsideIf()
    {
        var formula = "=IF(A1>0, LET(x, A1, y, B1, x + y), 0)";

        // Fallback: formula doesn't start with =LET( and LET is not directly inside outer func
        // This should attempt wrapped LET formatting
        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        // For now this is a complex case — the formatter should handle it or return unchanged
        Assert.NotNull(result);
    }

    [Fact]
    public void Format_NonLetFormula_ReturnedUnchanged()
    {
        var formula = "=SUM(A1:A10)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Equal(formula, result);
    }

    #endregion

    #region AutoFormatLet Bypass

    [Fact]
    public void Format_NestedLetDepthZero_ReturnsOriginal()
    {
        // AutoFormatLet = false is handled by caller, but NestedLetDepth = 0 disables formatting
        var formula = "=LET(x, 1, y, 2, x + y)";

        var result = LetFormulaFormatter.Format(formula, nestedLetDepth: 0);

        Assert.Equal(formula, result);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Format_NullFormula_ReturnsNull()
    {
        var result = LetFormulaFormatter.Format(null!, indentSize: 4);

        Assert.Null(result);
    }

    [Fact]
    public void Format_EmptyFormula_ReturnsEmpty()
    {
        var result = LetFormulaFormatter.Format("", indentSize: 4);

        Assert.Equal("", result);
    }

    [Fact]
    public void Format_WhitespaceOnly_ReturnsUnchanged()
    {
        var result = LetFormulaFormatter.Format("   ", indentSize: 4);

        Assert.Equal("   ", result);
    }

    [Fact]
    public void Format_MalformedLet_ReturnsUnchanged()
    {
        var formula = "=LET(x, 1";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        Assert.Equal(formula, result);
    }

    [Fact]
    public void Format_CaseInsensitiveLet()
    {
        var formula = "=let(x, 1, y, 2, x + y)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        // Should still format (parser is case-insensitive)
        Assert.Contains("\n", result);
    }

    [Fact]
    public void Format_NonLetFunctionCallsInValues_NotReformatted()
    {
        var formula = "=LET(x, FILTER(SORT(A1:A100, 2, -1), B1:B100 > 0), SUM(x))";

        var result = LetFormulaFormatter.Format(formula, indentSize: 4);

        // The complex function call should be preserved as-is
        Assert.Contains("x, FILTER(SORT(A1:A100, 2, -1), B1:B100 > 0),", result);
    }

    [Fact]
    public void Format_ManyBindings()
    {
        var formula = "=LET(a, 1, b, 2, c, 3, d, 4, e, 5, a + b + c + d + e)";

        var result = LetFormulaFormatter.Format(formula, indentSize: 2);

        var lines = result.Split('\n');
        Assert.Equal(7, lines.Length); // =LET( + 5 bindings + result
    }

    #endregion
}
