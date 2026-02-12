using FormulaBoss.Interception;
using FormulaBoss.Parsing;
using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class TranspilerTests
{
    private static TranspileResult Transpile(string source) => TranspileWithName(source, null);

    private static TranspileResult TranspileWithName(string source, string? preferredName)
    {
        var lexer = new Lexer(source);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens, source);
        var expression = parser.Parse();

        Assert.NotNull(expression);
        Assert.Empty(parser.Errors);

        var transpiler = new CSharpTranspiler();
        return transpiler.Transpile(expression, source, preferredName);
    }

    private static TranspileResult TranspileWithBindings(string source,
        Dictionary<string, ColumnBindingInfo> columnBindings)
    {
        var lexer = new Lexer(source);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens, source);
        var expression = parser.Parse();

        Assert.NotNull(expression);
        Assert.Empty(parser.Errors);

        var transpiler = new CSharpTranspiler();
        return transpiler.Transpile(expression, source, null, columnBindings);
    }

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

    [Fact]
    public void Transpiler_GeneratesCode_ForAdditionWithLiteral()
    {
        var result = Transpile("data.values.select(v => v + 1).toArray()");

        // When one side is a numeric literal, the lambda param should be cast
        Assert.Contains("Convert.ToDouble(v) + 1", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesCode_ForAdditionWithBothLambdaParams()
    {
        var result = Transpile("data.values.select(v => v + v).toArray()");

        // Both sides are lambda params but + can be string concatenation,
        // so we don't auto-cast (let C# infer or fail at compile time)
        Assert.Contains("(v + v)", result.SourceCode);
        Assert.DoesNotContain("Convert.ToDouble(v) + Convert.ToDouble(v)", result.SourceCode);
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
    public void Transpiler_Map_WithColorProperty_TranslatesCorrectly()
    {
        var result = Transpile("data.cells.map(c => c.color)");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("__MapPreserveShape__", result.SourceCode);
        // .color should translate to .Interior.ColorIndex
        Assert.Contains("Interior.ColorIndex", result.SourceCode);
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
        // Integer seed is converted to double to avoid type mismatch when lambda returns double
        Assert.Contains(".Aggregate(0d, (acc, x) =>", result.SourceCode);
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

    [Fact]
    public void Transpiler_Aggregate_ConvertsDecimalSeedToDouble()
    {
        // 0.0 in DSL is parsed as double 0, which becomes "0" in transpiler,
        // then converted to "0d" for type safety
        var result = Transpile("data.values.aggregate(0.0, (acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Aggregate(0d, (acc, x) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Aggregate_PreservesNonIntegerSeed()
    {
        // 1.5 in DSL has a fractional part, so it's preserved as "1.5"
        var result = Transpile("data.values.aggregate(1.5, (acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        // Non-integer double seeds are preserved as-is (no "d" suffix needed)
        Assert.Contains(".Aggregate(1.5, (acc, x) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Aggregate_ConvertsNegativeIntegerSeed()
    {
        var result = Transpile("data.values.aggregate(-1, (acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        // Negative integer seeds (parsed as unary minus) also get converted to double
        Assert.Contains(".Aggregate((-1d), (acc, x) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Aggregate_PreservesStringSeed()
    {
        var result = Transpile("data.values.aggregate(\"\", (acc, x) => acc + x)");

        Assert.False(result.RequiresObjectModel);
        // String seeds are preserved for string concatenation use cases
        Assert.Contains(".Aggregate(\"\", (acc, x) =>", result.SourceCode);
        // The + operator should NOT wrap in Convert.ToDouble for string concatenation
        Assert.Contains("(acc + x)", result.SourceCode);
        Assert.DoesNotContain("Convert.ToDouble(acc)", result.SourceCode);
    }

    #endregion

    #region Named Column Access

    [Fact]
    public void Transpiler_RowIndexAccess_WithColumnName_GeneratesHeaderLookup()
    {
        var result = Transpile("data.rows.where(r => r[Price] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__GetCol__(\"Price\")", result.SourceCode);
        Assert.Contains("__headers__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_RowMemberAccess_UsesTypedRows()
    {
        // Dot notation (r.Price) now triggers TypedRow generation
        var result = Transpile("data.rows.where(r => r.Price > 10).toArray()");

        Assert.True(result.RequiresObjectModel); // Dot notation triggers object model
        Assert.Contains("private class TypedRow_", result.SourceCode);
        Assert.Contains("public dynamic Price =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_RowReduce_WithNamedColumns_GeneratesHeaderDictionary()
    {
        var result = Transpile("data.rows.reduce(0, (acc, r) => acc + r[Price] * r[Qty])");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("Dictionary<string, int>", result.SourceCode);
        Assert.Contains("__GetCol__(\"Price\")", result.SourceCode);
        Assert.Contains("__GetCol__(\"Qty\")", result.SourceCode);
    }

    [Fact]
    public void Transpiler_WithHeaders_SetsHeaderContext()
    {
        var result = Transpile("data.withHeaders().rows.where(r => r[Name] == \"Test\").toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__headers__", result.SourceCode);
        Assert.Contains("__GetCol__(\"Name\")", result.SourceCode);
        // Header row should be skipped in __rows__ generation
        Assert.Contains("Enumerable.Range(1, rowCount - 1)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NumericIndex_StillWorksWithoutHeaders()
    {
        var result = Transpile("data.rows.where(r => r[0] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        // Numeric index should not generate header lookup
        Assert.DoesNotContain("__GetCol__", result.SourceCode);
        Assert.Contains("r[0]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_MixedAccess_SupportsNumericAndNamedColumns()
    {
        var result = Transpile("data.rows.select(r => r[0] + r[Total]).toArray()");

        Assert.False(result.RequiresObjectModel);
        // Should have header lookup for named column
        Assert.Contains("__GetCol__(\"Total\")", result.SourceCode);
        // And direct numeric access for index
        Assert.Contains("r[0]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ReduceAlias_WorksLikeAggregate()
    {
        var result = Transpile("data.rows.reduce(0, (acc, r) => acc + r[Amount])");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Aggregate(", result.SourceCode);
        Assert.Contains("__GetCol__(\"Amount\")", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NamedColumnAccess_GeneratesGetColHelper()
    {
        var result = Transpile("data.rows.where(r => r[Price] > 0).toArray()");

        Assert.False(result.RequiresObjectModel);
        // Should generate the __GetCol__ helper function
        Assert.Contains("Func<string, int> __GetCol__", result.SourceCode);
        // And the detailed error message pattern
        Assert.Contains("Column '{name}' not found", result.SourceCode);
    }

    [Fact]
    public void Transpiler_HeaderDictionary_IsCaseInsensitive()
    {
        var result = Transpile("data.rows.select(r => r[Price]).toArray()");

        Assert.False(result.RequiresObjectModel);
        // Should use case-insensitive comparison
        Assert.Contains("StringComparer.OrdinalIgnoreCase", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StringComparison_ConvertsColumnToString()
    {
        var result = Transpile("data.rows.where(r => r[Category] == \"Fruit\").toArray()");

        Assert.False(result.RequiresObjectModel);
        // Column access should be converted to string for comparison
        Assert.Contains("?.ToString()", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StringComparison_WithDotNotation()
    {
        // Dot notation triggers TypedRow which returns dynamic, so ToString() comparison still works
        var result = Transpile("data.rows.where(r => r.Status == \"Active\").toArray()");

        Assert.True(result.RequiresObjectModel); // Dot notation triggers object model
        Assert.Contains("private class TypedRow_", result.SourceCode);
        Assert.Contains("public dynamic Status =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ValuePath_NamedColumns_GeneratesHeaderFromFirstRow()
    {
        // Value-only path uses first row as headers
        var result = Transpile("data.rows.where(r => r[Price] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__headers__", result.SourceCode);
        // Value path builds headers from values[0, c]
        Assert.Contains("values[0, c]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NegativeIndex_AccessesFromEnd()
    {
        // r[-1] should access the last column
        var result = Transpile("data.rows.select(r => r[-1]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("r[r.Length - 1]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NegativeIndex_SecondFromEnd()
    {
        // r[-2] should access the second to last column
        var result = Transpile("data.rows.select(r => r[-2]).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("r[r.Length - 2]", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ColumnBindings_ResolvesBracketSyntax()
    {
        // When column bindings are provided, r[price] should resolve to r[__GetCol__(_price_colname_)]
        var columnBindings = new Dictionary<string, ColumnBindingInfo>
        {
            ["price"] = new("tblSales", "Price"),
            ["qty"] = new("tblSales", "Quantity")
        };
        var result = TranspileWithBindings("data.rows.reduce(0, (acc, r) => acc + r[price])", columnBindings);

        Assert.False(result.RequiresObjectModel);
        // Should use variable reference for dynamic column lookup
        Assert.Contains("__GetCol__(_price_colname_)", result.SourceCode);
        // Should generate UDF parameter for column name
        Assert.Contains("object price_col_param", result.SourceCode);
        // Should generate extractColName helper and use it
        Assert.Contains("extractColName", result.SourceCode);
        Assert.Contains("_price_colname_ = extractColName(price_col_param)", result.SourceCode);
        // Should track used column bindings
        Assert.NotNull(result.UsedColumnBindings);
        Assert.Contains("price", result.UsedColumnBindings);
    }

    [Fact]
    public void Transpiler_ColumnBindings_ResolvesDotSyntax()
    {
        // Dot notation with LET-bound column triggers TypedRow with pre-resolved index
        var columnBindings = new Dictionary<string, ColumnBindingInfo> { ["price"] = new("tblSales", "Price") };
        var result = TranspileWithBindings("data.rows.reduce(0, (acc, r) => acc + r.price)", columnBindings);

        Assert.True(result.RequiresObjectModel); // Dot notation triggers object model
        // TypedRow class with LET-bound property
        Assert.Contains("private class TypedRow_", result.SourceCode);
        Assert.Contains("public dynamic price =>", result.SourceCode);
        // Should track used column bindings
        Assert.NotNull(result.UsedColumnBindings);
        Assert.Contains("price", result.UsedColumnBindings);
    }

    [Fact]
    public void Transpiler_ColumnBindings_MultipleBindings()
    {
        // Multiple column bindings should all resolve correctly
        var columnBindings = new Dictionary<string, ColumnBindingInfo>
        {
            ["price"] = new("tblSales", "Price"),
            ["qty"] = new("tblSales", "Quantity")
        };
        var result = TranspileWithBindings("data.rows.reduce(0, (acc, r) => acc + r[price] * r[qty])", columnBindings);

        Assert.False(result.RequiresObjectModel);
        // Should use variable references for dynamic column lookup
        Assert.Contains("__GetCol__(_price_colname_)", result.SourceCode);
        Assert.Contains("__GetCol__(_qty_colname_)", result.SourceCode);
        // Should generate parameters for both column names
        Assert.Contains("object price_col_param", result.SourceCode);
        Assert.Contains("object qty_col_param", result.SourceCode);
        // Should track both used column bindings
        Assert.NotNull(result.UsedColumnBindings);
        Assert.Contains("price", result.UsedColumnBindings);
        Assert.Contains("qty", result.UsedColumnBindings);
    }

    #endregion

    #region Row Predicate Methods (find, some, every)

    [Fact]
    public void Transpiler_Find_GeneratesFirstOrDefault()
    {
        var result = Transpile("data.rows.find(r => r[0] > 10).toArray()");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".FirstOrDefault(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Some_GeneratesAny()
    {
        var result = Transpile("data.rows.some(r => r[0] > 10)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Any(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Every_GeneratesAll()
    {
        var result = Transpile("data.rows.every(r => r[0] > 0)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".All(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Find_WithNamedColumn_GeneratesHeaderLookup()
    {
        var result = Transpile("data.rows.find(r => r[Price] > 100)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".FirstOrDefault(", result.SourceCode);
        Assert.Contains("__GetCol__(\"Price\")", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Some_WithNamedColumn_GeneratesHeaderLookup()
    {
        var result = Transpile("data.rows.some(r => r[Price] > 100)");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains(".Any(", result.SourceCode);
        Assert.Contains("__GetCol__(\"Price\")", result.SourceCode);
    }

    #endregion

    #region Scan Method

    [Fact]
    public void Transpiler_Scan_GeneratesRunningReduction()
    {
        var result = Transpile("data.rows.scan(0, (acc, r) => acc + r[0])");

        Assert.False(result.RequiresObjectModel);
        // Should generate an IIFE that collects intermediate values
        Assert.Contains("var results = new List<object>();", result.SourceCode);
        Assert.Contains("foreach", result.SourceCode);
        Assert.Contains("results.Add(", result.SourceCode);
    }

    [Fact]
    public void Transpiler_Scan_WithNamedColumn_GeneratesHeaderLookup()
    {
        var result = Transpile("data.rows.scan(0, (sum, r) => sum + r[Amount])");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("__GetCol__(\"Amount\")", result.SourceCode);
        Assert.Contains("results.Add(", result.SourceCode);
    }

    #endregion

    #region Deep Property Access

    [Fact]
    public void Transpiler_DeepPropertyAccess_InteriorColorIndex()
    {
        var result = Transpile("data.cells.where(c => c.Interior.ColorIndex == 6).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(int)(c.Interior.ColorIndex ?? 0)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_InteriorPattern()
    {
        var result = Transpile("data.cells.where(c => c.Interior.Pattern == 1).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(int)(c.Interior.Pattern ?? 0)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_FontBold()
    {
        var result = Transpile("data.cells.where(c => c.Font.Bold == true).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(bool)(c.Font.Bold ?? false)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_FontSize()
    {
        var result = Transpile("data.cells.where(c => c.Font.Size > 12).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(double)(c.Font.Size ?? 11)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_FontName()
    {
        var result = Transpile("data.cells.select(c => c.Font.Name).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(string)(c.Font.Name ?? \"\")", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_CaseInsensitive()
    {
        // interior.colorindex should work same as Interior.ColorIndex
        var result = Transpile("data.cells.where(c => c.interior.colorindex == 6).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("ColorIndex", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DeepPropertyAccess_TriggersObjectModel()
    {
        // Accessing Interior or Font should trigger object model
        var result = Transpile("data.cells.where(c => c.Interior.Color > 0).toArray()");

        Assert.True(result.RequiresObjectModel);
    }

    #endregion

    #region Escape Hatch

    [Fact]
    public void Transpiler_EscapeHatch_BypassesValidation()
    {
        // @SomeCustomProperty should pass through without validation
        var result = Transpile("data.cells.where(c => c.@SomeCustomProperty != null).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("c.SomeCustomProperty", result.SourceCode);
        // Should NOT have null coalescing since we bypassed validation
        Assert.DoesNotContain("(int)(c.SomeCustomProperty", result.SourceCode);
    }

    [Fact]
    public void Transpiler_EscapeHatch_MixedWithValidated()
    {
        // Can mix escaped and validated properties
        var result = Transpile("data.cells.where(c => c.Interior.ColorIndex == 6 && c.@Custom == 1).toArray()");

        Assert.True(result.RequiresObjectModel);
        // Interior.ColorIndex should be properly cast
        Assert.Contains("(int)(c.Interior.ColorIndex ?? 0)", result.SourceCode);
        // Custom should be passed through verbatim
        Assert.Contains("c.Custom", result.SourceCode);
    }

    [Fact]
    public void Transpiler_EscapeHatch_DeepPath()
    {
        // Escape hatch on a deep property
        var result = Transpile("data.cells.where(c => c.@SomeObject.@SomeProperty == 1).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("c.SomeObject.SomeProperty", result.SourceCode);
    }

    #endregion

    #region Invalid Property Suggestions

    [Fact]
    public void Transpiler_InvalidProperty_ThrowsWithSuggestion_InteriorTypo()
    {
        var ex = Assert.Throws<TranspileException>(() =>
            Transpile("data.cells.where(c => c.Interior.Patern == 1).toArray()"));

        Assert.Contains("Patern", ex.Message);
        Assert.Contains("Pattern", ex.Message);
        Assert.Contains("Did you mean", ex.Message);
    }

    [Fact]
    public void Transpiler_InvalidProperty_ThrowsWithSuggestion_FontTypo()
    {
        var ex = Assert.Throws<TranspileException>(() =>
            Transpile("data.cells.where(c => c.Font.Bald == true).toArray()"));

        Assert.Contains("Bald", ex.Message);
        Assert.Contains("Bold", ex.Message);
    }

    [Fact]
    public void Transpiler_InvalidProperty_ThrowsForUnknown()
    {
        var ex = Assert.Throws<TranspileException>(() =>
            Transpile("data.cells.where(c => c.Interior.XYZ == 1).toArray()"));

        Assert.Contains("Unknown property 'XYZ' on Interior", ex.Message);
    }

    #endregion

    #region Null-Safe Access

    [Fact]
    public void Transpiler_SafeAccess_GeneratesTryCatch()
    {
        var result = Transpile("data.cells.select(c => c.@Comment?).toArray()");

        // Should generate try-catch wrapper that returns null on exception
        Assert.Contains("try { return c.Comment; }", result.SourceCode);
        Assert.Contains("catch { return null; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_SafeAccess_WithTypedProperty()
    {
        var result = Transpile("data.cells.select(c => c.Interior?).toArray()");

        // Should generate try-catch for typed property access
        Assert.Contains("try { return c.Interior; }", result.SourceCode);
        Assert.Contains("catch { return null; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NullCoalescing_PassesThrough()
    {
        var result = Transpile("data.cells.select(c => c.@Comment? ?? \"none\").toArray()");

        // Should pass through ?? operator to C#
        Assert.Contains("?? \"none\"", result.SourceCode);
    }

    [Fact]
    public void Transpiler_SafeAccess_MethodChain_WrapsEntireChain()
    {
        var result = Transpile("data.cells.select(c => c.@Comment?.Text() ?? \"none\").toArray()");

        // Should wrap the entire chain (c.Comment.Text()) in try-catch, not just c.Comment
        Assert.Contains("try { return c.Comment.Text(); }", result.SourceCode);
        Assert.Contains("catch { return null; }", result.SourceCode);
        Assert.Contains("?? \"none\"", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NullComparison_AutoWrapped()
    {
        var result = Transpile("data.cells.where(c => c.@Comment != null).toArray()");

        // Should wrap null comparison in try-catch for COM safety
        Assert.Contains("try { return c.Comment != null; }", result.SourceCode);
        Assert.Contains("catch { return false; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NullComparison_EqualNull_ReturnsTrue()
    {
        var result = Transpile("data.cells.where(c => c.@Comment == null).toArray()");

        // When checking == null, exception should return true (treating as null)
        Assert.Contains("try { return c.Comment == null; }", result.SourceCode);
        Assert.Contains("catch { return true; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_NullKeyword_TranspiledCorrectly()
    {
        var result = Transpile("data.cells.where(c => c.@Comment != null).toArray()");

        // "null" should be transpiled as the C# null keyword
        Assert.Contains("!= null", result.SourceCode);
    }

    #endregion

    #region Statement Lambda Tests

    [Fact]
    public void Transpiler_StatementLambda_ObjectModel_UsesDynamicParam()
    {
        var result = Transpile("data.cells.where(c => { return c.color == 6; })");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("(dynamic c) => { return c.color == 6; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_ValueOnly_UsesObjectParam()
    {
        var result = Transpile("data.values.where(v => { return Num(v) > 0; })");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("(object v) => { return Num(v) > 0; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_GeneratesHelperMethods()
    {
        var result = Transpile("data.values.where(v => { return Num(v) > 0; })");

        // Should generate all helper methods
        Assert.Contains("private static double Num(object x)", result.SourceCode);
        Assert.Contains("private static string Str(object x)", result.SourceCode);
        Assert.Contains("private static bool Bool(object x)", result.SourceCode);
        Assert.Contains("private static int Int(object x)", result.SourceCode);
        Assert.Contains("private static bool IsEmpty(object x)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_ExpressionLambda_DoesNotGenerateHelperMethods()
    {
        var result = Transpile("data.values.where(v => v > 0)");

        // Expression lambdas should NOT generate helper methods
        Assert.DoesNotContain("private static double Num(object x)", result.SourceCode);
        Assert.DoesNotContain("private static string Str(object x)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_MultiParam_UsesObjectTypes()
    {
        var result = Transpile("data.values.reduce((acc, x) => { return acc + x; })");

        Assert.False(result.RequiresObjectModel);
        Assert.Contains("(object acc, object x) => { return acc + x; }", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_WithRowContext()
    {
        var result = Transpile("data.rows.reduce(0, (acc, r) => { return Num(acc) + Num(r[0]); })");

        Assert.False(result.RequiresObjectModel);
        // Accumulator should be double (matches seed), row parameter should be object[]
        Assert.Contains("(double acc, object[] r) =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_PreservesBlockContent()
    {
        var result =
            Transpile("data.cells.where(c => { var color = (int)(c.Interior.ColorIndex ?? 0); return color == 6; })");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("var color = (int)(c.Interior.ColorIndex ?? 0); return color == 6;", result.SourceCode);
    }

    [Fact]
    public void Transpiler_HelperMethod_Num_HandlesNullAndNumbers()
    {
        var result = Transpile("data.values.where(v => { return Num(v) > 0; })");

        // Verify Num helper has correct implementation
        Assert.Contains("if (x == null) return 0;", result.SourceCode);
        Assert.Contains("if (x is double d) return d;", result.SourceCode);
    }

    [Fact]
    public void Transpiler_HelperMethod_IsEmpty_UsesReflection()
    {
        var result = Transpile("data.values.where(v => { return !IsEmpty(v); })");

        // IsEmpty should use reflection for ExcelEmpty check
        Assert.Contains("x.GetType().Name == \"ExcelEmpty\"", result.SourceCode);
    }

    #endregion

    #region Typed Row Objects

    [Fact]
    public void Transpiler_GeneratesTypedRow_WhenDotNotationUsed()
    {
        var result = Transpile("data.rows.where(r => r.Price > 10).toArray()");

        Assert.True(result.RequiresObjectModel); // Dot notation triggers object model
        Assert.Contains("private class TypedRow_", result.SourceCode);
        Assert.Contains("public dynamic Price =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesTypedRow_WithCellsProperty()
    {
        var result = Transpile("data.rows.select(r => r.cells[0].Value).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("private class TypedRow_", result.SourceCode);
        Assert.Contains("public readonly dynamic[] cells;", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesTypedRow_WithMultipleColumns()
    {
        var result = Transpile("data.rows.select(r => r.Price * r.Qty).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("public dynamic Price =>", result.SourceCode);
        Assert.Contains("public dynamic Qty =>", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesTypedRow_WithIndexers()
    {
        var result = Transpile("data.rows.select(r => r.Price).toArray()");

        Assert.Contains("public dynamic this[int index] => __values__[index];", result.SourceCode);
        Assert.Contains("public dynamic this[string colName] => __values__[_getCol(colName)];", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesRawRowsFallback_WhenTypedRowsUsed()
    {
        var result = Transpile("data.rows.select(r => r.Price).toArray()");

        Assert.Contains("var __rawRows__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_TranspilesRowsRows_ToRawRows_WhenTypedRowsUsed()
    {
        // When typed rows are used, data.rows.rows gives raw object[][] fallback
        var result = Transpile("data.rows.where(r => r.Price > 10).select(r => r).rows.count()");

        Assert.Contains("__rawRows__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_TranspilesRowsRows_ToRows_WhenNoTypedRows()
    {
        // When typed rows aren't used, data.rows.rows just stays as __rows__
        var result = Transpile("data.rows.rows.count()");

        // Should NOT generate TypedRow class
        Assert.DoesNotContain("private class TypedRow_", result.SourceCode);
        // data.rows.rows should become __rows__ (since __rawRows__ doesn't exist)
        Assert.Contains("__rows__", result.SourceCode);
    }

    [Fact]
    public void Transpiler_DoesNotGenerateTypedRow_WhenBracketNotationOnly()
    {
        var result = Transpile("data.rows.where(r => r[0] > 10).toArray()");

        // Bracket notation should NOT trigger typed rows (backward compat)
        Assert.DoesNotContain("private class TypedRow_", result.SourceCode);
    }

    [Fact]
    public void Transpiler_GeneratesTypedRow_ForLetBoundColumns()
    {
        var bindings = new Dictionary<string, ColumnBindingInfo>
        {
            ["p"] = new("tblSales", "Price"),
            ["q"] = new("tblSales", "Quantity")
        };

        var result = TranspileWithBindings("data.rows.select(r => r.p * r.q).toArray()", bindings);

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("private class TypedRow_", result.SourceCode);
        // LET-bound columns use pre-resolved indices
        Assert.Contains("public dynamic p =>", result.SourceCode);
        Assert.Contains("public dynamic q =>", result.SourceCode);
        Assert.Contains("_p_colIndex_", result.SourceCode);
        Assert.Contains("_q_colIndex_", result.SourceCode);
    }

    [Fact]
    public void Transpiler_StatementLambda_UsesTypedRowType()
    {
        var result = Transpile("data.rows.where(r => { return r.Price > 10; }).toArray()");

        Assert.True(result.RequiresObjectModel);
        Assert.Contains("private class TypedRow_", result.SourceCode);
        // Statement lambda should use TypedRow type for parameter
        Assert.Matches(@"\(TypedRow_\w+ r\)", result.SourceCode);
    }

    [Fact]
    public void Transpiler_SanitizesColumnNames_ForProperties()
    {
        // Column names with spaces or special chars should be sanitized
        var result = Transpile("data.rows.select(r => r.Sales_Amount).toArray()");

        Assert.Contains("public dynamic Sales_Amount =>", result.SourceCode);
    }

    #endregion
}
