using System.Text;

using FormulaBoss.Analysis;
using FormulaBoss.UI;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Verifies that reflection-based synthetic type generation produces correct output
///     matching the expected method signatures and type substitutions.
/// </summary>
public class TypedDocumentBuilderTests
{
    private static readonly WorkbookMetadata SingleTableMetadata = new(
        new[] { "Sales" },
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Sales"] = new[] { "Date", "Amount", "Region" }
        });

    [Fact]
    public void EmitTableTypes_RowCollection_HasAllExpectedMethods()
    {
        var sb = new StringBuilder();
        TypedDocumentBuilder.AppendUsings(sb);
        TypedDocumentBuilder.EmitTableTypes(sb, SingleTableMetadata);
        var source = sb.ToString();

        // RowCollection methods with correct type substitution
        Assert.Contains("class __SalesRowCollection : IEnumerable<__SalesRow>", source);
        Assert.Contains("Where(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("Select(Func<__SalesRow, object> selector)", source);
        Assert.Contains("Any(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("All(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("__SalesRow First(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("__SalesRow? FirstOrDefault(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("OrderBy(Func<__SalesRow, object> keySelector)", source);
        Assert.Contains("OrderByDescending(Func<__SalesRow, object> keySelector)", source);
        Assert.Contains("int Count()", source);
        Assert.Contains("Take(int count)", source);
        Assert.Contains("Skip(int count)", source);
        Assert.Contains("Distinct()", source);
        Assert.Contains("IExcelRange ToRange()", source);
        Assert.Contains("dynamic Aggregate(dynamic seed, Func<dynamic, __SalesRow, dynamic> func)", source);
        Assert.Contains("IExcelRange Scan(dynamic seed, Func<dynamic, __SalesRow, dynamic> func)", source);
        Assert.Contains("GroupBy(Func<__SalesRow, object> keySelector)", source);
        Assert.Contains("IEnumerator<__SalesRow> GetEnumerator()", source);
        Assert.Contains("IEnumerator IEnumerable.GetEnumerator()", source);
    }

    [Fact]
    public void EmitTableTypes_RowGroup_HasKeyAndInheritedMethods()
    {
        var sb = new StringBuilder();
        TypedDocumentBuilder.AppendUsings(sb);
        TypedDocumentBuilder.EmitTableTypes(sb, SingleTableMetadata);
        var source = sb.ToString();

        Assert.Contains("class __SalesRowGroup : IEnumerable<__SalesRow>", source);
        Assert.Contains("object? Key => default;", source);

        // Inherited RowCollection methods should be present
        Assert.Contains("__SalesRowCollection Where(Func<__SalesRow, bool> predicate)", source);
        Assert.Contains("__SalesRowCollection OrderBy(Func<__SalesRow, object> keySelector)", source);
        Assert.Contains("int Count()", source);
    }

    [Fact]
    public void EmitTableTypes_GroupedRowCollection_HasCorrectElementType()
    {
        var sb = new StringBuilder();
        TypedDocumentBuilder.AppendUsings(sb);
        TypedDocumentBuilder.EmitTableTypes(sb, SingleTableMetadata);
        var source = sb.ToString();

        Assert.Contains("class __SalesGroupedRowCollection : IEnumerable<__SalesRowGroup>", source);
        Assert.Contains("Select(Func<__SalesRowGroup, object> selector)", source);
        Assert.Contains("Where(Func<__SalesRowGroup, bool> predicate)", source);
        Assert.Contains("__SalesRowGroup First()", source);
        Assert.Contains("__SalesRowGroup First(Func<__SalesRowGroup, bool> predicate)", source);
    }

    [Fact]
    public void EmitTableTypes_Row_HasColumnProperties()
    {
        var sb = new StringBuilder();
        TypedDocumentBuilder.AppendUsings(sb);
        TypedDocumentBuilder.EmitTableTypes(sb, SingleTableMetadata);
        var source = sb.ToString();

        Assert.Contains("class __SalesRow : ExcelArray {", source);
        Assert.Contains("public ExcelScalar Date => default!;", source);
        Assert.Contains("public ExcelScalar Amount => default!;", source);
        Assert.Contains("public ExcelScalar Region => default!;", source);
        Assert.Contains("public ExcelScalar this[string columnName] => default!;", source);
        Assert.Contains("public int ColumnCount => 0;", source);
    }

    [Fact]
    public void EmitTableTypes_Table_InheritsExcelTable()
    {
        var sb = new StringBuilder();
        TypedDocumentBuilder.AppendUsings(sb);
        TypedDocumentBuilder.EmitTableTypes(sb, SingleTableMetadata);
        var source = sb.ToString();

        Assert.Contains("class __SalesTable : ExcelTable", source);
        Assert.Contains("new __SalesRowCollection Rows => default!;", source);
    }

    [Fact]
    public void EmitTableTypes_NullMetadata_ReturnsEmptyDictionary()
    {
        var sb = new StringBuilder();
        var result = TypedDocumentBuilder.EmitTableTypes(sb, null);

        Assert.Empty(result);
        Assert.Equal("", sb.ToString());
    }
}
