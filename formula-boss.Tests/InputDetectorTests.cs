using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class InputDetectorTests
{
    private readonly InputDetector _detector = new();

    #region Free Variable Detection (Parameters)

    [Fact]
    public void Detect_SingleFreeVariable()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => r[0] > 5)");

        Assert.Equal(["tbl"], result.Parameters);
    }

    [Fact]
    public void Detect_SingleFreeVariable_Count()
    {
        var result = _detector.Detect("tbl.Rows.Count()");

        Assert.Equal(["tbl"], result.Parameters);
    }

    [Fact]
    public void Detect_SingleFreeVariable_ChainedMethods()
    {
        var result = _detector.Detect("data.Rows.Where(r => (double)r[0] > 10).Select(r => r[1])");

        Assert.Equal(["data"], result.Parameters);
    }

    [Fact]
    public void Detect_MultipleFreeVariables()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] > minVal && (double)r[0] < maxVal)");

        Assert.Contains("tbl", result.Parameters);
        Assert.Contains("minVal", result.Parameters);
        Assert.Contains("maxVal", result.Parameters);
    }

    [Fact]
    public void Detect_NoFreeVariables_ThrowsOrEmpty()
    {
        // An expression with no free variables — all identifiers are lambda params or keywords
        // This should still work (could be a literal expression)
        var result = _detector.Detect("1 + 2");

        Assert.Empty(result.Parameters);
    }

    [Fact]
    public void Detect_LambdaParams_NotParameters()
    {
        // 'r' is a lambda param inside Where, not a free variable
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] > 5)");

        Assert.DoesNotContain("r", result.Parameters);
    }

    [Fact]
    public void Detect_FreeVariable_WithThreshold()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] < threshold)");

        Assert.Contains("tbl", result.Parameters);
        Assert.Contains("threshold", result.Parameters);
    }

    [Fact]
    public void Detect_ParametersAreSorted()
    {
        var result = _detector.Detect("z.Rows.Where(r => (double)r[0] > a && (double)r[0] < m)");

        Assert.Equal(["a", "m", "z"], result.Parameters);
    }

    [Fact]
    public void Detect_ForEachVariable_NotParameter()
    {
        var result = _detector.Detect("{ foreach (var country in countries) { } return countries; }");

        Assert.Equal(["countries"], result.Parameters);
        Assert.DoesNotContain("country", result.Parameters);
    }

    [Fact]
    public void Detect_PatternVariable_NotParameter()
    {
        var result = _detector.Detect("{ if (value is double d) return d; return 0; }");

        Assert.Equal(["value"], result.Parameters);
        Assert.DoesNotContain("d", result.Parameters);
    }

    [Fact]
    public void Detect_CatchVariable_NotParameter()
    {
        var result = _detector.Detect("{ try { return tbl.Sum(); } catch (System.Exception ex) { return 0; } }");

        Assert.Contains("tbl", result.Parameters);
        Assert.DoesNotContain("ex", result.Parameters);
    }

    #endregion

    #region Range Reference Detection

    [Fact]
    public void Detect_RangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("A1:B10.Rows.Count()");

        Assert.Single(result.Parameters);
        Assert.StartsWith("__range_", result.Parameters[0]);
        Assert.Contains("A1:B10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_AbsoluteRangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("$A$1:$B$10.Rows.Count()");

        Assert.Single(result.Parameters);
        Assert.StartsWith("__range_", result.Parameters[0]);
        Assert.Contains("$A$1:$B$10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_RangeRefNotTruncated()
    {
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

    #region Per-Variable Header Detection

    [Fact]
    public void Detect_StringBracketAccess_TracksVariable()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[\"Price\"] > 10)");

        Assert.Contains("tbl", result.HeaderVariables);
    }

    [Fact]
    public void Detect_NumericBracketAccess_NoHeaderVariable()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[0] > 10)");

        Assert.Empty(result.HeaderVariables);
    }

    [Fact]
    public void Detect_SpacedColumnName_TracksVariable()
    {
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[\"Unit Price\"] > 10)");

        Assert.Contains("tbl", result.HeaderVariables);
    }

    [Fact]
    public void Detect_MultipleVariables_OnlyHeaderOneTracked()
    {
        // tbl uses r["Col"] but threshold is just a scalar — only tbl in HeaderVariables
        var result = _detector.Detect("tbl.Rows.Where(r => (double)r[\"Price\"] > threshold)");

        Assert.Contains("tbl", result.HeaderVariables);
        Assert.DoesNotContain("threshold", result.HeaderVariables);
    }

    [Fact]
    public void Detect_CrossSheetRangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("Sheet2!A1:B10.Rows.Count()");

        Assert.Single(result.Parameters);
        Assert.StartsWith("__range_", result.Parameters[0]);
        Assert.Contains("Sheet2!A1:B10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_QuotedSheetRangeRef_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("'My Sheet'!A1:B10.Rows.Count()");

        Assert.Single(result.Parameters);
        Assert.StartsWith("__range_", result.Parameters[0]);
        Assert.Contains("'My Sheet'!A1:B10", result.RangeRefMap.Values);
    }

    [Fact]
    public void Detect_CrossSheetSingleCell_ReplacedWithPlaceholder()
    {
        var result = _detector.Detect("Sheet2!A1.Value");

        Assert.Single(result.Parameters);
        Assert.StartsWith("__range_", result.Parameters[0]);
        Assert.Contains("Sheet2!A1", result.RangeRefMap.Values);
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
        var (normalized, map) = InputDetector.PreprocessRangeRefs("tbl.Rows");

        Assert.Equal("tbl.Rows", normalized);
        Assert.Empty(map);
    }

    [Fact]
    public void PreprocessRangeRefs_SingleCellLikeVar_NotReplaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("A1.Rows");

        Assert.Equal("A1.Rows", normalized);
        Assert.Empty(map);
    }

    [Fact]
    public void PreprocessRangeRefs_CrossSheetRange_Replaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("Sheet2!A1:B10.Rows");

        Assert.DoesNotContain("!", normalized);
        Assert.Single(map);
        Assert.Contains("Sheet2!A1:B10", map.Values);
    }

    [Fact]
    public void PreprocessRangeRefs_QuotedSheetRange_Replaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("'My Sheet'!A1:B10.Rows");

        Assert.DoesNotContain("!", normalized);
        Assert.DoesNotContain("'", normalized);
        Assert.Single(map);
        Assert.Contains("'My Sheet'!A1:B10", map.Values);
    }

    [Fact]
    public void PreprocessRangeRefs_CrossSheetSingleCell_Replaced()
    {
        var (normalized, map) = InputDetector.PreprocessRangeRefs("Sheet2!A1.Value");

        Assert.DoesNotContain("!", normalized);
        Assert.Single(map);
        Assert.Contains("Sheet2!A1", map.Values);
    }

    #endregion

    #region Statement Block Detection

    [Fact]
    public void Detect_StatementBlock_DetectedCorrectly()
    {
        var result = _detector.Detect("{ var x = 1; return tbl.Sum() + x; }");

        Assert.True(result.IsStatementBlock);
        Assert.Equal(["tbl"], result.Parameters);
    }

    [Fact]
    public void Detect_StatementBlock_WithHeaderAccess()
    {
        var result = _detector.Detect("{ var count = tbl.Rows.Count(); return tbl.Rows.Where(r => r[\"Price\"] > 5).ToRange(); }");

        Assert.True(result.IsStatementBlock);
        Assert.Contains("tbl", result.Parameters);
        Assert.Contains("tbl", result.HeaderVariables);
    }

    [Fact]
    public void Detect_StatementBlock_MultipleParameters()
    {
        var result = _detector.Detect("{ if (threshold > 0) return tbl.Sum(); return otherTable.Sum(); }");

        Assert.True(result.IsStatementBlock);
        Assert.Equal(["otherTable", "tbl", "threshold"], result.Parameters);
    }

    [Fact]
    public void Detect_RegularExpression_NotStatementBlock()
    {
        var result = _detector.Detect("tbl.Sum()");

        Assert.False(result.IsStatementBlock);
    }

    [Fact]
    public void IsStatementBlock_LeadingBraceWithReturn_True()
    {
        Assert.True(InputDetector.IsStatementBlock("{ return 1; }"));
    }

    [Fact]
    public void IsStatementBlock_NoBrace_False()
    {
        Assert.False(InputDetector.IsStatementBlock("tbl.Sum()"));
    }

    #endregion
}
