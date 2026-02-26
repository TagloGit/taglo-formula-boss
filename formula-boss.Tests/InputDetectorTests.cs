using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class InputDetectorTests
{
    private readonly InputDetector _detector = new();

    #region Sugar Syntax Detection

    [Fact]
    public void Detect_SugarSyntax_SingleInput()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => r[0] > 5)");

        Assert.True(result.IsSugarSyntax);
        Assert.Equal(["tbl"], result.Inputs);
    }

    [Fact]
    public void Detect_SugarSyntax_Count()
    {
        var result = _detector.Detect("tbl.Rows.Count()");

        Assert.True(result.IsSugarSyntax);
        Assert.Equal(["tbl"], result.Inputs);
    }

    [Fact]
    public void Detect_SugarSyntax_ChainedMethods()
    {
        var result = _detector.Detect("data.Rows.Where(r => (double)r[0] > 10).Select(r => r[1])");

        Assert.True(result.IsSugarSyntax);
        Assert.Equal(["data"], result.Inputs);
    }

    #endregion

    #region Explicit Lambda Detection

    [Fact]
    public void Detect_ExplicitLambda_MultipleInputs()
    {
        var result = _detector.Detect("(tbl, maxVal) => tbl.Rows.Where(r => (double)r[0] < (double)maxVal)");

        Assert.False(result.IsSugarSyntax);
        Assert.Equal(["tbl", "maxVal"], result.Inputs);
    }

    [Fact]
    public void Detect_ExplicitLambda_SingleInput()
    {
        var result = _detector.Detect("(tbl) => tbl.Rows.Count()");

        Assert.False(result.IsSugarSyntax);
        Assert.Equal(["tbl"], result.Inputs);
    }

    [Fact]
    public void Detect_SimpleLambda_SingleInput()
    {
        var result = _detector.Detect("tbl => tbl.Rows.Count()");

        Assert.False(result.IsSugarSyntax);
        Assert.Equal(["tbl"], result.Inputs);
    }

    #endregion

    #region Statement Body Detection

    [Fact]
    public void Detect_StatementBody_Detected()
    {
        var result = _detector.Detect("(tbl) => { var c = tbl.Rows.Count(); return c; }");

        Assert.True(result.IsStatementBody);
        Assert.Equal(["tbl"], result.Inputs);
    }

    [Fact]
    public void Detect_ExpressionBody_NotStatement()
    {
        var result = _detector.Detect("(tbl) => tbl.Rows.Count()");

        Assert.False(result.IsStatementBody);
    }

    #endregion

    #region Range Reference Detection

    [Fact]
    public void Detect_RangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("A1:B10.Rows.Count()");

        Assert.True(result.IsSugarSyntax);
        Assert.Single(result.Inputs);
        Assert.StartsWith("__range_", result.Inputs[0]);
        Assert.Contains("A1:B10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_AbsoluteRangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("$A$1:$B$10.Rows.Count()");

        Assert.Single(result.Inputs);
        Assert.StartsWith("__range_", result.Inputs[0]);
        Assert.Contains("$A$1:$B$10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_RangeRefNotTruncated()
    {
        // Ensure "A1:A10" doesn't get truncated to just "A1"
        var result = _detector.Detect("A1:A10.Rows.Count()");

        var rangeRef = result.RangeRefMap.Values.Single();
        Assert.Equal("A1:A10", rangeRef);
    }

    #endregion

    #region Object Model Detection

    [Fact]
    public void Detect_CellAccess_RequiresObjectModel()
    {
        var result = _detector.Detect("tbl.Rows.Select(r => r[0].Cell.Color)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Detect_CellsAccess_RequiresObjectModel()
    {
        var result = _detector.Detect("tbl.Cells.Where(c => c.Color == 6)");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Detect_NoCellAccess_DoesNotRequireObjectModel()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] > 5)");

        Assert.False(result.RequiresObjectModel);
    }

    #endregion

    #region Free Variable Detection

    [Fact]
    public void Detect_FreeVariable_Detected()
    {
        var result = _detector.Detect("(tbl) => tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("threshold", result.FreeVariables);
    }

    [Fact]
    public void Detect_NoFreeVariables_WhenAllBound()
    {
        var result = _detector.Detect("(tbl) => tbl.Rows.Count()");

        Assert.Empty(result.FreeVariables);
    }

    [Fact]
    public void Detect_MultipleFreeVariables()
    {
        var result = _detector.Detect("(tbl) => tbl.Rows.Where(r => (double)r[0] > minVal && (double)r[0] < maxVal)");

        Assert.Contains("minVal", result.FreeVariables);
        Assert.Contains("maxVal", result.FreeVariables);
    }

    [Fact]
    public void Detect_LambdaParams_NotFreeVariables()
    {
        // 'r' is a lambda param inside Where, not a free variable
        var result = _detector.Detect("(tbl) => tbl.Rows.Where(r => (double)r[0] > 5)");

        Assert.DoesNotContain("r", result.FreeVariables);
    }

    #endregion

    #region String Bracket Access

    [Fact]
    public void Detect_StringBracketAccess_Detected()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[\"Price\"] > 10)");

        Assert.True(result.HasStringBracketAccess);
    }

    [Fact]
    public void Detect_NumericBracketAccess_NotStringBracket()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] > 10)");

        Assert.False(result.HasStringBracketAccess);
    }

    [Fact]
    public void Detect_SpacedColumnName_Detected()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[\"Unit Price\"] > 10)");

        Assert.True(result.HasStringBracketAccess);
    }

    #endregion

    #region Range Ref Preprocessing

    [Fact]
    public void PreprocessRangeRefs_ColonRange_Replaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("A1:B10.Rows");

        Assert.DoesNotContain(":", normalized);
        Assert.Single(map);
        Assert.Contains("A1:B10", map.Values);
    }

    [Fact]
    public void PreprocessRangeRefs_AbsoluteRange_Replaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("$A$1:$B$10.Rows");

        Assert.DoesNotContain("$", normalized);
        Assert.Single(map);
        Assert.Contains("$A$1:$B$10", map.Values);
    }

    [Fact]
    public void PreprocessRangeRefs_SimpleIdentifier_NotReplaced()
    {
        // "tbl" should not be treated as a range ref even though it has letters
        var (normalized, map) = InputDetector.PreprocessRangeRefs("tbl.Rows");

        Assert.Equal("tbl.Rows", normalized);
        Assert.Empty(map);
    }

    [Fact]
    public void PreprocessRangeRefs_SingleCellLikeVar_NotReplaced()
    {
        // A1 without colon or $ should be treated as a regular identifier
        var (normalized, map) = InputDetector.PreprocessRangeRefs("A1.Rows");

        Assert.Equal("A1.Rows", normalized);
        Assert.Empty(map);
    }

    #endregion
}
