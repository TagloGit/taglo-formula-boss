using FormulaBoss.Parsing;
using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class TranspilerTests
{
    #region Object Model Detection

    [Fact]
    public void Transpiler_DetectsObjectModel_WhenCellsUsed()
    {
        var result = Transpile("data.cells.toArray()");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Transpiler_DetectsObjectModel_WhenColorUsed()
    {
        var result = Transpile("data.cells.where(c => c.color == 6).toArray()");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Transpiler_DetectsObjectModel_WhenRowUsed()
    {
        var result = Transpile("data.cells.select(c => c.row).toArray()");

        Assert.True(result.RequiresObjectModel);
    }

    [Fact]
    public void Transpiler_DetectsValueOnly_WhenValuesUsed()
    {
        var result = Transpile("data.values.toArray()");

        Assert.False(result.RequiresObjectModel);
    }

    [Fact]
    public void Transpiler_DetectsValueOnly_WhenNoObjectModelProperties()
    {
        var result = Transpile("data.values.where(v => v > 0).toArray()");

        Assert.False(result.RequiresObjectModel);
    }

    #endregion

    #region Range Reference Support

    [Fact]
    public void Transpiler_HandlesRangeReference_ValueOnly()
    {
        var result = Transpile("A1:B10.values.where(v => v > 0).toArray()");

        Assert.False(result.RequiresObjectModel);
        // Range reference should become __source__ which becomes values in value-only path
        Assert.Contains("__values__.Where", result.SourceCode);
    }

    [Fact]
    public void Transpiler_HandlesRangeReference_WithObjectModel()
    {
        var result = Transpile("A1:J10.cells.where(c => c.color == 6).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("__cells__.Where", result.SourceCode);
    }

    [Fact]
    public void Transpiler_HandlesAbsoluteRangeReference()
    {
        var result = Transpile("$A$1:$B$10.values.sum()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__values__", result.SourceCode);
    }

    #endregion

    #region Method Name Generation

    [Fact]
    public void Transpiler_GeneratesConsistentMethodName()
    {
        var result1 = Transpile("data.values.toArray()");
        var result2 = Transpile("data.values.toArray()");

        Assert.Equal(result1.MethodName, result2.MethodName);
    }

    [Fact]
    public void Transpiler_GeneratesDifferentMethodNames_ForDifferentExpressions()
    {
        var result1 = Transpile("data.values.toArray()");
        var result2 = Transpile("data.cells.toArray()");

        Assert.NotEqual(result1.MethodName, result2.MethodName);
    }

    [Fact]
    public void Transpiler_MethodNameStartsWithUdfPrefix()
    {
        var result = Transpile("data.values.toArray()");

        Assert.StartsWith("__udf_", result.MethodName);
    }

    [Fact]
    public void Transpiler_UsesPreferredName_WhenProvided()
    {
        var result = TranspileWithName("data.values.toArray()", "myCustomUdf");

        Assert.Equal("MYCUSTOMUDF", result.MethodName);
    }

    [Fact]
    public void Transpiler_SanitizesPreferredName_ToUppercase()
    {
        var result = TranspileWithName("data.values.toArray()", "coloredCells");

        Assert.Equal("COLOREDCELLS", result.MethodName);
    }

    [Fact]
    public void Transpiler_SanitizesPreferredName_RemovesInvalidChars()
    {
        var result = TranspileWithName("data.values.toArray()", "my-udf!@#name");

        Assert.Equal("MYUDFNAME", result.MethodName);
    }

    [Fact]
    public void Transpiler_SanitizesPreferredName_AddsUnderscoreIfStartsWithDigit()
    {
        var result = TranspileWithName("data.values.toArray()", "123abc");

        Assert.Equal("_123ABC", result.MethodName);
    }

    [Fact]
    public void Transpiler_FallsBackToHash_WhenPreferredNameIsEmpty()
    {
        var result = TranspileWithName("data.values.toArray()", "");

        Assert.StartsWith("__udf_", result.MethodName);
    }

    [Fact]
    public void Transpiler_FallsBackToHash_WhenPreferredNameIsWhitespace()
    {
        var result = TranspileWithName("data.values.toArray()", "   ");

        Assert.StartsWith("__udf_", result.MethodName);
    }

    #endregion

    #region Code Generation - Literals

    [Fact]
    public void Transpiler_GeneratesCode_ForNumberLiteral()
    {
        var result = Transpile("data.values.where(v => v > 42).toArray()");

        Assert.Contains("42", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForStringLiteral()
    {
        var result = Transpile("data.values.where(v => v == \"hello\").toArray()");

        Assert.Contains("\"hello\"", result.SourceCode);
    }

    #endregion

    #region Code Generation - Operators

    [Fact]
    public void Transpiler_GeneratesCode_ForComparisonOperators()
    {
        var result = Transpile("data.values.where(v => v >= 10).toArray()");

        Assert.Contains(">=", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForLogicalOperators()
    {
        var result = Transpile("data.values.where(v => v > 0 && v < 100).toArray()");

        Assert.Contains("&&", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForArithmeticOperators()
    {
        var result = Transpile("data.values.select(v => v * 2 + 1).toArray()");

        Assert.Contains("*", result.SourceCode);
        Assert.Contains("+", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForArithmeticWithCast()
    {
        var result = Transpile("data.values.select(v => v * 2).toArray()");

        // Lambda parameter should be cast to double for arithmetic
        Assert.Contains("Convert.ToDouble(v) * 2", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForArithmeticWithBothSidesCast()
    {
        var result = Transpile("data.values.select(v => v * v).toArray()");

        // Both sides should be cast when both are lambda params
        Assert.Contains("Convert.ToDouble(v) * Convert.ToDouble(v)", result.SourceCode);
    }

    #endregion

    #region Code Generation - LINQ Methods

    [Fact]
    public void Transpiler_GeneratesCode_ForWhere()
    {
        var result = Transpile("data.values.where(v => v > 0).toArray()");

        Assert.Contains(".Where(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForSelect()
    {
        var result = Transpile("data.values.select(v => v * 2).toArray()");

        Assert.Contains(".Select(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForToArray()
    {
        var result = Transpile("data.values.toArray()");

        Assert.Contains(".ToArray()", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForTakePositive()
    {
        var result = Transpile("data.values.take(5).toArray()");

        Assert.Contains(".Take(5)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForTakeNegative()
    {
        var result = Transpile("data.values.take(-2).toArray()");

        // Negative take should generate TakeLast
        Assert.Contains(".TakeLast(2)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForSkipPositive()
    {
        var result = Transpile("data.values.skip(3).toArray()");

        Assert.Contains(".Skip(3)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForSkipNegative()
    {
        var result = Transpile("data.values.skip(-2).toArray()");

        // Negative skip should generate SkipLast
        Assert.Contains(".SkipLast(2)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForChainedMethods()
    {
        var result = Transpile("data.values.where(v => v > 0).select(v => v * 2).toArray()");

        Assert.Contains(".Where(", result.SourceCode);
        Assert.Contains(".Select(", result.SourceCode);
        Assert.Contains(".ToArray()", result.SourceCode);
    }

    #endregion

    #region Code Generation - Cell Properties

    [Fact]
    public void Transpiler_GeneratesCode_ForCellValue()
    {
        var result = Transpile("data.cells.select(c => c.value).toArray()");

        Assert.Contains(".Value", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForCellColor()
    {
        var result = Transpile("data.cells.where(c => c.color == 6).toArray()");

        Assert.Contains("Interior.ColorIndex", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForCellRow()
    {
        var result = Transpile("data.cells.select(c => c.row).toArray()");

        Assert.Contains(".Row", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForCellCol()
    {
        var result = Transpile("data.cells.select(c => c.col).toArray()");

        Assert.Contains(".Column", result.SourceCode);
    }

    #endregion

    #region Code Generation - Structure

    [Fact]
    public void Transpiler_GeneratesCode_WithoutExcelDnaAttributes()
    {
        // We don't generate ExcelDNA attributes in the code anymore due to assembly identity issues.
        // Registration is handled manually via RegisterDelegates in DynamicCompiler.
        var result = Transpile("data.values.toArray()");

        Assert.DoesNotContain("[ExcelFunction(", result.SourceCode);
        Assert.DoesNotContain("[ExcelArgument(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_WithCorrectMethodName()
    {
        var result = Transpile("data.values.toArray()");

        Assert.Contains($"public static object {result.MethodName}", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_WithUsingStatements()
    {
        var result = Transpile("data.values.toArray()");

        Assert.Contains("using System;", result.SourceCode);
        Assert.Contains("using System.Linq;", result.SourceCode);
        // Note: We avoid using ExcelDna.Integration due to assembly identity mismatch issues
        Assert.DoesNotContain("using ExcelDna.Integration;", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ObjectModel_UsesInlineReflection()
    {
        var result = Transpile("data.cells.toArray()");

        // Object model path uses inline reflection for ExcelReference → Range conversion
        Assert.Contains("ExcelDnaUtil", result.SourceCode);
        // Uses property-based address building instead of XlCall
        Assert.Contains("RowFirst", result.SourceCode);
        Assert.Contains("ColumnFirst", result.SourceCode);
        Assert.Contains("dynamic range", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ValueOnly_DoesNotIncludeInterop()
    {
        var result = Transpile("data.values.toArray()");

        // Value-only path shouldn't reference Excel interop for cell access
        Assert.DoesNotContain("range.Cast<Microsoft.Office.Interop.Excel.Range>()", result.SourceCode);
    }

    #endregion

    #region Full Expression Tests

    [Fact]
    public void Transpiler_FullExpression_FilterByColor()
    {
        var result = Transpile("data.cells.where(c => c.color == 6).select(c => c.value).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains(".Where(", result.SourceCode);
        Assert.Contains("Interior.ColorIndex", result.SourceCode);
        Assert.Contains(".Select(", result.SourceCode);
        Assert.Contains(".Value", result.SourceCode);
        Assert.Contains(".ToArray()", result.SourceCode);
    }

    [Fact]
    public void Transpiler_FullExpression_FilterPositiveValues()
    {
        var result = Transpile("data.values.where(v => v > 0).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Where(v => (Convert.ToDouble(v) > 0))", result.SourceCode);
        Assert.Contains(".ToArray()", result.SourceCode);
    }

    #endregion

    #region Rows and Cols Support

    [Fact]
    public void Transpiler_GeneratesCode_ForRows()
    {
        var result = Transpile("data.rows.toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__rows__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForCols()
    {
        var result = Transpile("data.cols.toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__cols__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForRowsWhere_ValueOnly()
    {
        var result = Transpile("data.rows.where(r => r[0] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__rows__.Where(", result.SourceCode);
        Assert.Contains("r[0]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForColsOrderBy()
    {
        var result = Transpile("data.cols.orderBy(c => c[0]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__cols__.OrderBy(", result.SourceCode);
        Assert.Contains("c[0]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForRowsTake()
    {
        var result = Transpile("data.rows.take(3).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__rows__.Take(3)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForRowsTakeLast()
    {
        var result = Transpile("data.rows.take(-2).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__rows__.TakeLast(2)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_RowsDoesNotRequireObjectModel()
    {
        var result = Transpile("data.rows.where(r => r[0] > 0).toArray()");

        Assert.False(result.RequiresObjectModel);
    }

    #endregion

    #region Map Support

    [Fact]
    public void Transpiler_GeneratesCode_ForMap()
    {
        var result = Transpile("data.map(v => v * 2)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__MapPreserveShape__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForMap_WithObjectModel()
    {
        var result = Transpile("data.cells.map(c => c.value * 2)");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("__MapPreserveShape__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_MapPreservesShapeHelper_IsGenerated()
    {
        var result = Transpile("data.map(v => v * 2)");

        // Verify the helper function is generated
        Assert.Contains("Func<Func<object, object>, object[,]> __MapPreserveShape__", result.SourceCode);
    }

    #endregion

    #region Index Access

    [Fact]
    public void Transpiler_GeneratesCode_ForIndexAccess()
    {
        var result = Transpile("data.rows.select(r => r[0]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("r[0]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForMultipleIndexAccess()
    {
        var result = Transpile("data.rows.where(r => r[0] > r[1]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("r[0]", result.SourceCode);
        Assert.Contains("r[1]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_IndexAccess_CastsToDoubleForComparison()
    {
        var result = Transpile("data.rows.where(r => r[0] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        // r[0] should be cast to double for comparison with numeric literal
        Assert.Contains("Convert.ToDouble(r[0])", result.SourceCode);
    }

    [Fact]
    public void Transpiler_IndexAccess_CastsToDoubleForArithmetic()
    {
        var result = Transpile("data.rows.select(r => r[0] * 2).toArray()");

        Assert.False(result.RequiresObjectModel);
        // r[0] should be cast to double for arithmetic
        Assert.Contains("Convert.ToDouble(r[0])", result.SourceCode);
    }

    #endregion

    #region GroupBy Support

    [Fact]
    public void Transpiler_GeneratesCode_ForGroupBy()
    {
        var result = Transpile("data.rows.groupBy(r => r[0]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".GroupBy(", result.SourceCode);
        Assert.Contains(".SelectMany(g => g)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GroupBy_FlattensResult()
    {
        var result = Transpile("data.values.groupBy(v => v).toArray()");

        Assert.False(result.RequiresObjectModel);
        // GroupBy should be followed by SelectMany to flatten
        Assert.Contains("GroupBy(v => v).SelectMany(g => g)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GroupBy_WithAggregator_ReturnsKeyValuePairs()
    {
        var result = Transpile("data.rows.groupBy(r => r[0], g => g.count()).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".GroupBy(", result.SourceCode);
        Assert.Contains("g.Key", result.SourceCode);
        Assert.Contains("g.Count()", result.SourceCode);
        Assert.Contains("new object[]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GroupBy_WithSumAggregator()
    {
        var result = Transpile("data.rows.groupBy(r => r[0], g => g.sum(r => r[1])).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".GroupBy(", result.SourceCode);
        Assert.Contains("g.Key", result.SourceCode);
        Assert.Contains(".Sum(", result.SourceCode);
    }

    #endregion

    #region Aggregate Support

    [Fact]
    public void Transpiler_GeneratesCode_ForAggregate_WithSeed()
    {
        var result = Transpile("data.values.aggregate(0, (acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Aggregate(0, (acc, x) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForAggregate_NoSeed()
    {
        var result = Transpile("data.values.aggregate((acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Aggregate((acc, x) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Aggregate_HandlesMultiParamLambda()
    {
        var result = Transpile("data.rows.aggregate((acc, r) => acc + r[0])");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("(acc, r) =>", result.SourceCode);
    }

    #endregion

    private static TranspileResult Transpile(string source)
    {
        return TranspileWithName(source, null);
    }

    private static TranspileResult TranspileWithName(string source, string? preferredName)
    {
        var lexer = new Lexer(source);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);
        var expression = parser.Parse();

        Assert.NotNull(expression);
        Assert.Empty(parser.Errors);

        var transpiler = new CSharpTranspiler();
        return transpiler.Transpile(expression, source, preferredName);
    }
}
