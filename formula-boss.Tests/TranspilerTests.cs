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

    private static TranspileResult Transpile(string source)
    {
        var lexer = new Lexer(source);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);
        var expression = parser.Parse();

        Assert.NotNull(expression);
        Assert.Empty(parser.Errors);

        var transpiler = new CSharpTranspiler();
        return transpiler.Transpile(expression, source);
    }
}
