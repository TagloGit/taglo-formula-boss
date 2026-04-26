using FormulaBoss.Analysis;
using FormulaBoss.UI;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using Xunit;

namespace FormulaBoss.Tests;

public class SemanticAnalysisServiceTests
{
    private static readonly WorkbookMetadata TwoTableMetadata = new(
        new[] { "Sales", "Products" },
        new[] { "TaxRate" },
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Sales"] = new[] { "Date", "Amount", "Region" },
            ["Products"] = new[] { "Name", "Price", "Category" }
        });

    private readonly SemanticAnalysisService _service = new();

    [Fact]
    public void BuildSemanticModel_SimpleExpression_ReturnsResult()
    {
        var result = _service.BuildSemanticModel(
            "Sales.Rows.Count()", false, TwoTableMetadata, null);

        Assert.NotNull(result);
        Assert.NotNull(result.SemanticModel);
    }

    [Fact]
    public void BuildSemanticModel_NullMetadata_ReturnsResult()
    {
        var result = _service.BuildSemanticModel("42", false, null, null);

        Assert.NotNull(result);
    }

    [Fact]
    public void BuildSemanticModel_StatementBlock_ReturnsResult()
    {
        var result = _service.BuildSemanticModel(
            "{ var x = 42; return x; }", true, TwoTableMetadata, null);

        Assert.NotNull(result);
    }

    [Fact]
    public void BuildSemanticModel_WithLetBindings_DeclaresVariables()
    {
        var bindings = new[] { ("t", "Sales") };
        var result = _service.BuildSemanticModel(
            "t.Rows.Count()", false, TwoTableMetadata, bindings);

        Assert.NotNull(result);
        // The expression should compile without errors for the table access
        var source = result.SemanticModel.SyntaxTree.ToString();
        Assert.Contains("__SalesTable t = default!", source);
    }

    [Fact]
    public void GetTypeAtOffset_TableVariable_ReturnsTableType()
    {
        var result = _service.BuildSemanticModel(
            "Sales", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        // Offset 0 is the start of "Sales"
        var type = _service.GetTypeAtOffset(result, 0);
        Assert.NotNull(type);
        Assert.Contains("Sales", type.Name);
    }

    [Fact]
    public void GetTypeAtOffset_VarDeclaration_ReturnsInferredType()
    {
        var result = _service.BuildSemanticModel(
            "{ var x = 42; return x; }", true, TwoTableMetadata, null);
        Assert.NotNull(result);

        // Find "var" in the expression - offset 2 (after "{ ")
        var type = _service.GetTypeAtOffset(result, 2);
        Assert.NotNull(type);
        Assert.Equal("Int32", type.Name);
    }

    [Fact]
    public void GetTypeAtOffset_LambdaParameter_ReturnsRowType()
    {
        var result = _service.BuildSemanticModel(
            "Sales.Rows.Where(r => r.Amount > 0)", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        // Find "r" after "Where(" - the lambda parameter
        var expr = "Sales.Rows.Where(r => r.Amount > 0)";
        var rOffset = expr.IndexOf("r =>", StringComparison.Ordinal);
        var type = _service.GetTypeAtOffset(result, rOffset);
        Assert.NotNull(type);
        Assert.Contains("Row", type.Name);
    }

    [Fact]
    public void GetTypeAtOffset_NamedRange_ReturnsExcelArray()
    {
        var result = _service.BuildSemanticModel(
            "TaxRate", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        var type = _service.GetTypeAtOffset(result, 0);
        Assert.NotNull(type);
        Assert.Equal("ExcelArray", type.Name);
    }

    [Fact]
    public void FormatTypeForDisplay_SyntheticTableType_FormatsAsExcelTable()
    {
        var result = _service.BuildSemanticModel(
            "Sales", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        var type = _service.GetTypeAtOffset(result, 0);
        Assert.NotNull(type);

        var display = _service.FormatTypeForDisplay(type, TwoTableMetadata);
        Assert.Equal("ExcelTable (Sales)", display);
    }

    [Fact]
    public void FormatTypeForDisplay_SyntheticRowType_FormatsWithColumns()
    {
        var result = _service.BuildSemanticModel(
            "Sales.Rows.First(r => true)", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        // Get the type of the whole expression (should be the row type)
        _service.GetTypeAtOffset(result, 0);
        // The expression starts with "Sales" which is the table, get the result type
        // by looking at the full expression result — use a simpler approach
        var resultForExpr = _service.BuildSemanticModel(
            "Sales.Rows.First(r => true)", false, TwoTableMetadata, null);
        Assert.NotNull(resultForExpr);

        // Check the 'r' parameter type display
        var rResult = _service.BuildSemanticModel(
            "Sales.Rows.Where(r => r.Amount > 0)", false, TwoTableMetadata, null);
        Assert.NotNull(rResult);
        var rOffset = "Sales.Rows.Where(".Length;
        var rType = _service.GetTypeAtOffset(rResult, rOffset);
        Assert.NotNull(rType);

        var rDisplay = _service.FormatTypeForDisplay(rType, TwoTableMetadata);
        Assert.Equal("Row {Date, Amount, Region}", rDisplay);
    }

    [Fact]
    public void FormatTypeForDisplay_SyntheticRowCollection_FormatsAsRowCollection()
    {
        var result = _service.BuildSemanticModel(
            "Sales.Rows", false, TwoTableMetadata, null);
        Assert.NotNull(result);

        // "Sales.Rows" — get type at "Rows" position
        var roResult = _service.BuildSemanticModel(
            "Sales.Rows.Where(r => true)", false, TwoTableMetadata, null);
        Assert.NotNull(roResult);

        // The Where result is a RowCollection
        // Get the full expression type by evaluating a simpler expression
        var rowCollResult = _service.BuildSemanticModel(
            "Sales.Rows.Where(r => true)", false, TwoTableMetadata, null);
        Assert.NotNull(rowCollResult);

        // The expression "Sales.Rows" is a row collection
        var rowsResult = _service.BuildSemanticModel(
            "Sales.Rows", false, TwoTableMetadata, null);
        Assert.NotNull(rowsResult);

        // Get type at start of expression (Sales) then check Rows
        // Actually, let's just verify FormatTypeForDisplay with a known synthetic name
        // by getting the type of "Sales.Rows" expression
        // Offset for "Rows" in "Sales.Rows" is 6
        var rowsType = _service.GetTypeAtOffset(rowsResult, 6);
        Assert.NotNull(rowsType);

        var display = _service.FormatTypeForDisplay(rowsType, TwoTableMetadata);
        Assert.Equal("RowCollection (Sales)", display);
    }

    [Fact]
    public void FormatTypeForDisplay_ExcelScalar_UnchangedName()
    {
        // Use the bracket indexer on Row to obtain ExcelScalar — column names are not
        // emitted as dot-access properties on the synthetic Row.
        var display = _service.FormatTypeForDisplay(
            GetTypeForExpression("Sales.Rows.First(r => true)[\"Amount\"]"),
            TwoTableMetadata);
        Assert.Equal("ExcelScalar", display);
    }

    [Fact]
    public void FormatTypeForDisplay_PlainType_UnchangedName()
    {
        var display = _service.FormatTypeForDisplay(
            GetTypeForExpression("42"),
            TwoTableMetadata);
        Assert.Equal("int", display);
    }

    [Fact]
    public void FormatTypeForDisplay_RowWithManyColumns_Truncates()
    {
        var metadata = new WorkbookMetadata(
            new[] { "BigTable" },
            Array.Empty<string>(),
            new Dictionary<string, IReadOnlyList<string>> { ["BigTable"] = new[] { "A", "B", "C", "D", "E", "F" } });

        var result = _service.BuildSemanticModel(
            "BigTable.Rows.Where(r => true)", false, metadata, null);
        Assert.NotNull(result);

        var rOffset = "BigTable.Rows.Where(".Length;
        var rType = _service.GetTypeAtOffset(result, rOffset);
        Assert.NotNull(rType);

        var display = _service.FormatTypeForDisplay(rType, metadata);
        Assert.Equal("Row {A, B, C, D, ...}", display);
    }

    [Fact]
    public void BuildSemanticModel_LetBinding_ResolvesNamedRange()
    {
        var bindings = new[] { ("r", "TaxRate") };
        var result = _service.BuildSemanticModel(
            "r", false, TwoTableMetadata, bindings);
        Assert.NotNull(result);

        var type = _service.GetTypeAtOffset(result, 0);
        Assert.NotNull(type);
        Assert.Equal("ExcelArray", type.Name);
    }

    private ITypeSymbol GetTypeForExpression(string expression)
    {
        var result = _service.BuildSemanticModel(expression, false, TwoTableMetadata, null);
        Assert.NotNull(result);

        // Get type at the end of the expression (the full expression's result type)
        // We need the type of the entire expression — find __result's type
        var root = result.SemanticModel.SyntaxTree.GetRoot();
        var varDecl = root.DescendantNodes()
            .OfType<VariableDeclaratorSyntax>()
            .FirstOrDefault(v => v.Identifier.Text == "__result");
        Assert.NotNull(varDecl);

        var symbol = result.SemanticModel.GetDeclaredSymbol(varDecl);
        Assert.NotNull(symbol);
        return ((ILocalSymbol)symbol).Type;
    }
}
