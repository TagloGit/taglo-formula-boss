using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaParserTests
{
    #region IsLetFormula Tests

    [Fact]
    public void IsLetFormula_ReturnsFalse_ForNull()
    {
        Assert.False(LetFormulaParser.IsLetFormula(null));
    }

    [Fact]
    public void IsLetFormula_ReturnsFalse_ForEmptyString()
    {
        Assert.False(LetFormulaParser.IsLetFormula(""));
    }

    [Fact]
    public void IsLetFormula_ReturnsFalse_ForNonLetFormula()
    {
        Assert.False(LetFormulaParser.IsLetFormula("=SUM(A1:A10)"));
    }

    [Fact]
    public void IsLetFormula_ReturnsTrue_ForSimpleLet()
    {
        Assert.True(LetFormulaParser.IsLetFormula("=LET(x, 1, x)"));
    }

    [Fact]
    public void IsLetFormula_ReturnsTrue_CaseInsensitive()
    {
        Assert.True(LetFormulaParser.IsLetFormula("=let(x, 1, x)"));
        Assert.True(LetFormulaParser.IsLetFormula("=Let(x, 1, x)"));
    }

    [Fact]
    public void IsLetFormula_ReturnsFalse_ForLetInString()
    {
        // "LET" appearing in a string literal shouldn't match
        Assert.False(LetFormulaParser.IsLetFormula("=\"LET(x, 1, x)\""));
    }

    #endregion

    #region TryParse Simple Cases

    [Fact]
    public void TryParse_ParsesSimpleLet()
    {
        var formula = "=LET(x, 1, x)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Single(structure.Bindings);
        Assert.Equal("x", structure.Bindings[0].VariableName);
        Assert.Equal("1", structure.Bindings[0].Value);
        Assert.False(structure.Bindings[0].HasBacktick);
        Assert.Equal("x", structure.ResultExpression);
    }

    [Fact]
    public void TryParse_ParsesMultipleBindings()
    {
        var formula = "=LET(a, 1, b, 2, a + b)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(2, structure.Bindings.Count);
        Assert.Equal("a", structure.Bindings[0].VariableName);
        Assert.Equal("1", structure.Bindings[0].Value);
        Assert.Equal("b", structure.Bindings[1].VariableName);
        Assert.Equal("2", structure.Bindings[1].Value);
        Assert.Equal("a + b", structure.ResultExpression);
    }

    [Fact]
    public void TryParse_DetectsBacktickInBinding()
    {
        var formula = "=LET(data, A1:A10, filtered, `data.where(v => v > 0)`, SUM(filtered))";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(2, structure.Bindings.Count);
        Assert.False(structure.Bindings[0].HasBacktick);
        Assert.True(structure.Bindings[1].HasBacktick);
        Assert.Equal("`data.where(v => v > 0)`", structure.Bindings[1].Value);
    }

    #endregion

    #region TryParse Nested Structures

    [Fact]
    public void TryParse_HandlesNestedParentheses()
    {
        var formula = "=LET(x, SUM(A1:A5), y, AVERAGE(B1:B5), x + y)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(2, structure.Bindings.Count);
        Assert.Equal("SUM(A1:A5)", structure.Bindings[0].Value);
        Assert.Equal("AVERAGE(B1:B5)", structure.Bindings[1].Value);
    }

    [Fact]
    public void TryParse_HandlesStringLiteralsWithCommas()
    {
        var formula = "=LET(msg, \"hello, world\", msg)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Single(structure.Bindings);
        Assert.Equal("\"hello, world\"", structure.Bindings[0].Value);
    }

    [Fact]
    public void TryParse_HandlesNestedFunctionsWithCommas()
    {
        var formula = "=LET(result, IF(A1>0, B1, C1), result * 2)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Single(structure.Bindings);
        Assert.Equal("IF(A1>0, B1, C1)", structure.Bindings[0].Value);
    }

    [Fact]
    public void TryParse_HandlesMultilineFormula()
    {
        var formula = @"=LET(data, A1:F20,
     coloredCells, `data.cells.where(c => c.color != -4142)`,
     result)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(2, structure.Bindings.Count);
        Assert.Equal("data", structure.Bindings[0].VariableName);
        Assert.Equal("coloredCells", structure.Bindings[1].VariableName.Trim());
        Assert.True(structure.Bindings[1].HasBacktick);
    }

    [Fact]
    public void TryParse_PreservesOriginalFormula()
    {
        var formula = "=LET(x, 1, x)";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(formula, structure.OriginalFormula);
    }

    #endregion

    #region TryParse Error Cases

    [Fact]
    public void TryParse_ReturnsFalse_ForNonLetFormula()
    {
        var success = LetFormulaParser.TryParse("=SUM(A1:A10)", out var structure);

        Assert.False(success);
        Assert.Null(structure);
    }

    [Fact]
    public void TryParse_ReturnsFalse_ForMalformedLet()
    {
        // Missing closing paren
        var success = LetFormulaParser.TryParse("=LET(x, 1, x", out var structure);

        Assert.False(success);
        Assert.Null(structure);
    }

    [Fact]
    public void TryParse_ReturnsFalse_ForTooFewArguments()
    {
        // LET needs at least 3 arguments (name, value, result)
        var success = LetFormulaParser.TryParse("=LET(x, 1)", out var structure);

        Assert.False(success);
        Assert.Null(structure);
    }

    [Fact]
    public void TryParse_ReturnsFalse_ForEvenNumberOfArguments()
    {
        // Even number of arguments means missing result expression
        var success = LetFormulaParser.TryParse("=LET(x, 1, y, 2)", out var structure);

        Assert.False(success);
        Assert.Null(structure);
    }

    #endregion

    #region ExtractBacktickExpression Tests

    [Fact]
    public void ExtractBacktickExpression_ReturnsNull_ForNoBackticks()
    {
        var result = LetFormulaParser.ExtractBacktickExpression("SUM(A1:A10)");

        Assert.Null(result);
    }

    [Fact]
    public void ExtractBacktickExpression_ExtractsExpression()
    {
        var result = LetFormulaParser.ExtractBacktickExpression("`data.where(v => v > 0)`");

        Assert.Equal("data.where(v => v > 0)", result);
    }

    [Fact]
    public void ExtractBacktickExpression_ReturnsNull_ForUnclosedBacktick()
    {
        var result = LetFormulaParser.ExtractBacktickExpression("`data.where(v => v > 0)");

        Assert.Null(result);
    }

    #endregion

    #region Complex Real-World Scenarios

    [Fact]
    public void TryParse_HandlesFullFeatureExample()
    {
        var formula = @"=LET(data, A1:F20,
     coloredCells, `data.cells.where(c => c.color != -4142)`,
     result, `coloredCells.select(c => c.value * 2).toArray()`,
     SUM(result))";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(3, structure.Bindings.Count);

        Assert.Equal("data", structure.Bindings[0].VariableName);
        Assert.Equal("A1:F20", structure.Bindings[0].Value);
        Assert.False(structure.Bindings[0].HasBacktick);

        Assert.True(structure.Bindings[1].HasBacktick);
        Assert.True(structure.Bindings[2].HasBacktick);

        Assert.Equal("SUM(result)", structure.ResultExpression.Trim());
    }

    [Fact]
    public void TryParse_HandlesMixedBacktickAndNormalBindings()
    {
        var formula = "=LET(raw, A1:A100, threshold, 50, filtered, `raw.where(v => v > threshold)`, COUNT(filtered))";

        var success = LetFormulaParser.TryParse(formula, out var structure);

        Assert.True(success);
        Assert.NotNull(structure);
        Assert.Equal(3, structure.Bindings.Count);
        Assert.False(structure.Bindings[0].HasBacktick); // raw
        Assert.False(structure.Bindings[1].HasBacktick); // threshold
        Assert.True(structure.Bindings[2].HasBacktick);  // filtered
    }

    #endregion
}
