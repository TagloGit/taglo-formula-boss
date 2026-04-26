using FormulaBoss.Compilation;
using FormulaBoss.UI;
using FormulaBoss.UI.Completion;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

using Xunit;

namespace FormulaBoss.Tests;

public class SyntheticDocumentBuilderTests
{
    private static readonly WorkbookMetadata TwoTableMetadata = new(
        new[] { "Sales", "Products" },
        new[] { "TaxRate" },
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Sales"] = new[] { "Date", "Amount", "Region" },
            ["Products"] = new[] { "Name", "Price", "Category" }
        });

    [Fact]
    public void Build_ContainsUsings()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("using System;", source);
        Assert.Contains("using System.Linq;", source);
        Assert.Contains("using FormulaBoss.Runtime;", source);
    }

    [Fact]
    public void Build_TypedRowClass_DoesNotEmitColumnNameProperties()
    {
        // Column names are surfaced by CompletionHelpers.BuildRowCompletions as bracket-
        // inserting items; emitting them as synthetic properties on Row would duplicate
        // that path and force the completion provider to strip them out again.
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("class __SalesRow", source);
        Assert.DoesNotContain("public ExcelScalar Date", source);
        Assert.DoesNotContain("public ExcelScalar Amount", source);
        Assert.DoesNotContain("public ExcelScalar Region", source);
    }

    [Fact]
    public void Build_GeneratesTypedRowCollection_WithMethods()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("class __SalesRowCollection", source);
        Assert.Contains("Where(Func<__SalesRow, bool>", source);
        Assert.Contains("Select(Func<__SalesRow, object>", source);
        Assert.Contains("OrderBy(Func<__SalesRow, object>", source);
    }

    [Fact]
    public void Build_GeneratesTypedTableClass()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("class __SalesTable : ExcelTable", source);
        Assert.Contains("new __SalesRowCollection Rows", source);
    }

    [Fact]
    public void Build_GeneratesTypedTableClass_WithColumnIndexer()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("new Column this[string columnName]", source);
    }

    [Fact]
    public void Build_GeneratesMultipleTables()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("class __SalesRow", source);
        Assert.Contains("class __ProductsRow", source);
        Assert.Contains("class __SalesTable", source);
        Assert.Contains("class __ProductsTable", source);
    }

    [Fact]
    public void Build_DeclaresLetBindingVariables()
    {
        var formula = "=LET(t, Sales, `t.rows.where(r => r.`)";
        var textUp = "=LET(t, Sales, `t.rows.where(r => r.";

        var (source, _) = SyntheticDocumentBuilder.Build(formula, textUp, TwoTableMetadata);
        Assert.Contains("__SalesTable t = default!", source);
    }

    [Fact]
    public void Build_EmbedsDslExpression()
    {
        var formula = "=LET(t, Sales, `t.rows.where(r => r.`)";
        var textUp = "=LET(t, Sales, `t.rows.where(r => r.";

        var (source, _) = SyntheticDocumentBuilder.Build(formula, textUp, TwoTableMetadata);
        Assert.Contains("t.rows.where(r => r.", source);
    }

    [Fact]
    public void Build_CaretOffset_IsAtEndOfExpression()
    {
        var formula = "=`Sales.`";
        var textUp = "=`Sales.";

        var (source, offset) = SyntheticDocumentBuilder.Build(formula, textUp, TwoTableMetadata);
        // The caret should be at the end of "Sales." within the synthetic source
        Assert.True(offset > 0);
        Assert.True(offset <= source.Length);
        // The character before the caret should be '.'
        Assert.Equal('.', source[offset - 1]);
    }

    [Fact]
    public void Build_HandlesNullMetadata()
    {
        var (source, offset) = SyntheticDocumentBuilder.Build("=`x.`", "=`x.", null);
        Assert.NotEmpty(source);
        Assert.True(offset > 0);
    }

    [Fact]
    public void Build_StatementBlock_EmitsLocalFunction()
    {
        var formula = "=LET(myFunc, `{ var s = new StringBuilder(); s. }`)";
        var textUp = "=LET(myFunc, `{ var s = new StringBuilder(); s.";

        var (source, _) = SyntheticDocumentBuilder.Build(formula, textUp, TwoTableMetadata);
        Assert.Contains("object __userBlock()", source);
        Assert.Contains("var s = new StringBuilder();", source);
        Assert.DoesNotContain("var __result = {", source);
    }

    [Fact]
    public void BuildForDiagnostics_StatementBlock_EmitsLocalFunction()
    {
        var formula = "=LET(myFunc, `{ var s = new StringBuilder(); return s.ToString(); }`)";

        var result = SyntheticDocumentBuilder.BuildForDiagnostics(formula, TwoTableMetadata);
        Assert.NotNull(result);
        Assert.Contains("object __userBlock()", result.Source);
        Assert.Contains("var s = new StringBuilder();", result.Source);
        Assert.DoesNotContain("var __result = {", result.Source);
    }

    [Fact]
    public void Build_NamedRange_DeclaredAsExcelArray()
    {
        var formula = "=LET(r, TaxRate, `r.`)";
        var textUp = "=LET(r, TaxRate, `r.";

        var (source, _) = SyntheticDocumentBuilder.Build(formula, textUp, TwoTableMetadata);
        Assert.Contains("ExcelArray r = default!", source);
    }

    [Fact]
    public void Build_TypedRowClass_InheritsExcelArray()
    {
        var (source, _) = SyntheticDocumentBuilder.Build("=`Sales.`", "=`Sales.", TwoTableMetadata);
        Assert.Contains("class __SalesRow : ExcelArray", source);
    }

    [Fact]
    public void BuildForDiagnostics_LinqOnRow_CompilesWithoutDiagnostics()
    {
        // Mirrors the issue's repro: LINQ-style methods on a Row resolved via .First()
        // should be valid C# in the synthetic document.
        var formula = "=`Sales.Rows.First().Skip(1).FirstOrDefault(p => p != null)`";

        AssertSyntheticDiagnosticFree(formula, TwoTableMetadata);
    }

    [Fact]
    public void BuildForDiagnostics_ExcelArrayMethodsOnRow_CompilesWithoutDiagnostics()
    {
        // ExcelArray surface (Where, Map, Skip, Take, Count) on a Row.
        var formula = "=`Sales.Rows.First().Skip(1).Where(s => s != null).Map(s => s).Take(2).Count()`";

        AssertSyntheticDiagnosticFree(formula, TwoTableMetadata);
    }

    private static void AssertSyntheticDiagnosticFree(string formula, WorkbookMetadata metadata)
    {
        var result = SyntheticDocumentBuilder.BuildForDiagnostics(formula, metadata);
        Assert.NotNull(result);

        var syntaxTree = CSharpSyntaxTree.ParseText(result.Source);
        var compilation = CSharpCompilation.Create(
            "DiagnosticCheck",
            new[] { syntaxTree },
            MetadataReferenceProvider.GetMetadataReferences(),
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        // Only check diagnostics that fall within the embedded expression — errors elsewhere
        // (e.g. ambiguous overloads in the synthetic stubs) aren't what this test is verifying.
        var exprStart = result.ExpressionStartInSynthetic;
        var exprEnd = exprStart + result.ExpressionLength;
        var errors = compilation.GetDiagnostics()
            .Where(d => d.Severity == DiagnosticSeverity.Error)
            .Where(d =>
            {
                var span = d.Location.SourceSpan;
                return span.Start >= exprStart && span.End <= exprEnd;
            })
            .Select(d => d.ToString())
            .ToList();

        Assert.True(errors.Count == 0,
            $"Expected no diagnostic errors on embedded expression but got:\n{string.Join("\n", errors)}\n\nSource:\n{result.Source}");
    }
}
