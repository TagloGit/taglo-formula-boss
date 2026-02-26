using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class InputDetectorTests
{
    // === Single-input sugar syntax ===

    [Fact]
    public void Sugar_SimpleMethodChain_ExtractsPrimaryInput()
    {
        var result = InputDetector.Detect("tblCountries.Rows.Where(r => r.Population > 1000)");

        Assert.Single(result.Inputs);
        Assert.Equal("tblCountries", result.Inputs[0]);
        Assert.False(result.IsExplicitLambda);
    }

    [Fact]
    public void Sugar_BodyIsEntireExpression()
    {
        var expr = "tblCountries.Rows.Where(r => r.Population > 1000)";
        var result = InputDetector.Detect(expr);

        Assert.Equal(expr, result.Body);
    }

    [Fact]
    public void Sugar_NoObjectModel_WhenNoCellAccess()
    {
        var result = InputDetector.Detect("tblCountries.Rows.Where(r => r.Population > 1000)");

        Assert.False(result.RequiresObjectModel);
    }

    [Fact]
    public void Sugar_DetectsCellUsage()
    {
        var result = InputDetector.Detect("tblCountries.Rows.Where(r => r.Population.Cell.Color == 6)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Sugar_DetectsCellsUsage()
    {
        var result = InputDetector.Detect("tblCountries.Cells.Where(c => c.Color == 6)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Sugar_DetectsFreeVariables_WithKnownLetVars()
    {
        var knownVars = new HashSet<string> { "maxPop", "pConts" };
        var result = InputDetector.Detect(
            "tblCountries.Rows.Where(r => r.Population < maxPop && pConts.Any(c => c == r.Continent))",
            knownVars);

        Assert.Contains("maxPop", result.FreeVariables);
        Assert.Contains("pConts", result.FreeVariables);
    }

    [Fact]
    public void Sugar_FreeVariables_FilteredToKnownLetVars()
    {
        var knownVars = new HashSet<string> { "maxPop" };
        var result = InputDetector.Detect(
            "tblCountries.Rows.Where(r => r.Population < maxPop && unknownThing > 0)",
            knownVars);

        Assert.Single(result.FreeVariables);
        Assert.Equal("maxPop", result.FreeVariables[0]);
    }

    [Fact]
    public void Sugar_IsNotStatementBody()
    {
        var result = InputDetector.Detect("tblCountries.Rows.Count()");

        Assert.False(result.IsStatementBody);
    }

    // === Explicit lambda syntax ===

    [Fact]
    public void Lambda_SingleParam_ExtractsInput()
    {
        var result = InputDetector.Detect("(tbl) => tbl.Rows.Where(r => r.X > 0)");

        Assert.Single(result.Inputs);
        Assert.Equal("tbl", result.Inputs[0]);
        Assert.True(result.IsExplicitLambda);
    }

    [Fact]
    public void Lambda_MultiParam_ExtractsAllInputs()
    {
        var result = InputDetector.Detect("(tbl, maxPop) => tbl.Rows.Where(r => r.Population < maxPop)");

        Assert.Equal(2, result.Inputs.Count);
        Assert.Equal("tbl", result.Inputs[0]);
        Assert.Equal("maxPop", result.Inputs[1]);
        Assert.True(result.IsExplicitLambda);
    }

    [Fact]
    public void Lambda_BodyExtracted_ExcludesArrow()
    {
        var result = InputDetector.Detect("(tbl) => tbl.Rows.Count()");

        Assert.Equal("tbl.Rows.Count()", result.Body);
    }

    [Fact]
    public void Lambda_StatementBlock_Detected()
    {
        var result = InputDetector.Detect("(tbl) => { return tbl.Rows.Count(); }");

        Assert.True(result.IsStatementBody);
        Assert.Contains("return tbl.Rows.Count();", result.Body);
    }

    [Fact]
    public void Lambda_MultiParam_StatementBlock()
    {
        var result = InputDetector.Detect(
            "(tbl, other) => { if (tbl.Rows.Count() > 0) return \"yes\"; return \"no\"; }");

        Assert.Equal(2, result.Inputs.Count);
        Assert.True(result.IsStatementBody);
    }

    [Fact]
    public void Lambda_DetectsCellUsage()
    {
        var result = InputDetector.Detect("(tbl) => tbl.Rows.Where(r => r.Price.Cell.Color == 6)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Lambda_DetectsFreeVariables()
    {
        var knownVars = new HashSet<string> { "threshold" };
        var result = InputDetector.Detect(
            "(tbl) => tbl.Rows.Where(r => r.Value > threshold)",
            knownVars);

        Assert.Contains("threshold", result.FreeVariables);
    }

    [Fact]
    public void Lambda_ParamsNotCountedAsFreeVars()
    {
        var knownVars = new HashSet<string> { "tbl", "maxPop" };
        var result = InputDetector.Detect(
            "(tbl, maxPop) => tbl.Rows.Where(r => r.Population < maxPop)",
            knownVars);

        Assert.Empty(result.FreeVariables);
    }

    [Fact]
    public void Lambda_InnerLambdaParamsNotFree()
    {
        var result = InputDetector.Detect(
            "(tbl) => tbl.Rows.Where(r => r.X > 0).Select(r => r.Y)");

        // 'r' is a lambda param, not a free variable
        Assert.DoesNotContain("r", result.FreeVariables);
    }

    // === SimpleLambda (no parens) ===

    [Fact]
    public void SimpleLambda_SingleParam()
    {
        var result = InputDetector.Detect("tbl => tbl.Rows.Count()");

        Assert.Single(result.Inputs);
        Assert.Equal("tbl", result.Inputs[0]);
        Assert.True(result.IsExplicitLambda);
    }
}
