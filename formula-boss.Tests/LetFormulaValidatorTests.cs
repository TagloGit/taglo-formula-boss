using FormulaBoss.Interception;

using Xunit;

namespace FormulaBoss.Tests;

public class LetFormulaValidatorTests
{
    [Fact]
    public void Validate_ReturnsNoErrors_ForValidLetFormula()
    {
        var errors = LetFormulaValidator.Validate("=LET(x, 1, y, 2, x+y)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForSingleBinding()
    {
        var errors = LetFormulaValidator.Validate("=LET(x, 5, x)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForNonLetFormula()
    {
        var errors = LetFormulaValidator.Validate("=SUM(A1:A10)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForEmptyString()
    {
        var errors = LetFormulaValidator.Validate("");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForIncompleteFormulaWhileTyping()
    {
        // No closing paren — user is still typing
        var errors = LetFormulaValidator.Validate("=LET(x, 5, ");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForIncompleteLetJustOpened()
    {
        var errors = LetFormulaValidator.Validate("=LET(");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_DetectsMissingResultExpression()
    {
        // Even number of args: x, 1, y, 2 — no result
        var errors = LetFormulaValidator.Validate("=LET(x, 1, y, 2)");
        Assert.Single(errors);
        Assert.Contains("Missing result", errors[0].Message);
    }

    [Fact]
    public void Validate_DetectsTooFewArguments_OnlyName()
    {
        var errors = LetFormulaValidator.Validate("=LET(x)");
        Assert.Single(errors);
        Assert.Contains("requires at least one binding", errors[0].Message);
    }

    [Fact]
    public void Validate_DetectsMissingResult_NameAndValueOnly()
    {
        var errors = LetFormulaValidator.Validate("=LET(x, 5)");
        Assert.Single(errors);
        Assert.Contains("Missing result", errors[0].Message);
    }

    [Fact]
    public void Validate_DetectsEmptyLet()
    {
        var errors = LetFormulaValidator.Validate("=LET()");
        Assert.Single(errors);
        Assert.Contains("empty", errors[0].Message);
    }

    [Fact]
    public void Validate_DetectsUnbalancedParentheses()
    {
        // Inner SUM( is not closed, but there's a trailing ) that closes LET incorrectly
        var errors = LetFormulaValidator.Validate("=LET(x, SUM(1,2, x)");
        Assert.Single(errors);
        Assert.Contains("Unbalanced parentheses", errors[0].Message);
    }

    [Fact]
    public void Validate_NoError_ForUnbalancedParensWhileTyping()
    {
        // No closing paren at all — user is still typing
        var errors = LetFormulaValidator.Validate("=LET(x, SUM(1,2");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_DetectsInvalidVariableName_WithSpaces()
    {
        var errors = LetFormulaValidator.Validate("=LET(x y, 5, x)");
        Assert.Single(errors);
        Assert.Contains("Invalid variable name", errors[0].Message);
        Assert.Contains("x y", errors[0].Message);
    }

    [Fact]
    public void Validate_DetectsInvalidVariableName_StartsWithNumber()
    {
        var errors = LetFormulaValidator.Validate("=LET(1x, 5, 1x)");
        Assert.Single(errors);
        Assert.Contains("Invalid variable name", errors[0].Message);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForBacktickExpressions()
    {
        var errors = LetFormulaValidator.Validate("=LET(t, Table1, `t.Rows.Count()`)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ReturnsNoErrors_ForNestedFunctions()
    {
        var errors = LetFormulaValidator.Validate("=LET(x, SUM(1,2,3), y, AVERAGE(4,5), x+y)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_ErrorPositionsAreCorrect_MissingResult()
    {
        var formula = "=LET(x, 1, y, 2)";
        var errors = LetFormulaValidator.Validate(formula);
        Assert.Single(errors);
        // The last argument " 2" should be highlighted
        var error = errors[0];
        Assert.True(error.StartOffset > 0);
        Assert.True(error.Length > 0);
    }

    [Fact]
    public void Validate_ErrorPositionsAreCorrect_InvalidName()
    {
        var formula = "=LET(a b, 5, x)";
        var errors = LetFormulaValidator.Validate(formula);
        Assert.Single(errors);
        var error = errors[0];
        // "a b" starts at index 5
        Assert.Equal(5, error.StartOffset);
        Assert.Equal(3, error.Length);
    }

    [Fact]
    public void Validate_CaseInsensitive()
    {
        var errors = LetFormulaValidator.Validate("=let(x, 1, x)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_MultipleErrors()
    {
        // Invalid name AND missing result
        var errors = LetFormulaValidator.Validate("=LET(1x, 5, 2y, 10)");
        Assert.True(errors.Count >= 2);
    }

    [Fact]
    public void Validate_AllowsUnderscoreInVariableNames()
    {
        var errors = LetFormulaValidator.Validate("=LET(_x, 5, my_var, 10, _x)");
        Assert.Empty(errors);
    }

    [Fact]
    public void Validate_AllowsDotInVariableNames()
    {
        // Excel LET allows dotted names like sheet1.range
        var errors = LetFormulaValidator.Validate("=LET(a.b, 5, a.b)");
        Assert.Empty(errors);
    }
}
