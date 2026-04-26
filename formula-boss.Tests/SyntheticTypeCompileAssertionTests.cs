using FormulaBoss.Compilation;
using FormulaBoss.UI;
using FormulaBoss.UI.Completion;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Compile-asserts that the synthetic types emitted by <c>TypedDocumentBuilder</c>
///     give the editor a faithful view of the runtime API for representative usage patterns.
///     Complements <see cref="SyntheticMemberSyncTests" />, which only proves every runtime
///     member is tagged for emission. Here we run the editor's Roslyn compilation on documented
///     patterns and assert zero errors on the embedded expression — adding a new pattern is
///     a one-line addition, so these tests double as living documentation of what the editor
///     promises end-to-end.
/// </summary>
public class SyntheticTypeCompileAssertionTests
{
    private static readonly WorkbookMetadata Metadata = new(
        new[] { "Players" },
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Players"] = new[] { "Player", "Item", "Value" }
        });

    [Fact]
    public void Row_LinqAndCellEscalation_Compiles()
    {
        // Row inherits ExcelArray, so .Skip → IExcelRange, then .FirstOrDefault(Func<ExcelScalar,bool>)
        // resolves on the inherited surface. ExcelScalar.Cell escalates to a Cell with .Bold.
        AssertCompiles("Players.Rows.First().Skip(1).FirstOrDefault(c => c.Cell.Bold)");
    }

    [Fact]
    public void Row_MapWithPrimitiveTResult_Compiles()
    {
        // ExcelArray.Map<TResult>(Func<ExcelScalar, TResult>) — TResult=double via implicit
        // ExcelValue→double conversion. Spec 0013 documents this pattern.
        AssertCompiles("Players.Rows.First().Map(c => (double)c * 2)");
    }

    [Fact]
    public void Row_StringIndexer_Compiles()
    {
        AssertCompiles("Players.Rows.First()[\"Player\"]");
    }

    [Fact]
    public void Row_IntIndexer_Compiles()
    {
        // Inherited from ExcelArray — synthetic Row does not redeclare it.
        AssertCompiles("Players.Rows.First()[0]");
    }

    [Fact]
    public void Table_StringIndexer_OnColumn_Compiles()
    {
        // Table[string] → Column (which extends ExcelArray, so .Sum() resolves).
        AssertCompiles("Players[\"Player\"].Sum()");
    }

    [Fact]
    public void GroupedRowCollection_FirstAndKey_Compiles()
    {
        // RowCollection.GroupBy → __PlayersGroupedRowCollection, .First() → __PlayersRowGroup,
        // .Key inherited via [SyntheticMember] on RowGroup.
        AssertCompiles("Players.Rows.GroupBy(r => r[\"Player\"]).First().Key");
    }

    private static void AssertCompiles(string innerExpression)
    {
        var formula = $"=`{innerExpression}`";
        var result = SyntheticDocumentBuilder.BuildForDiagnostics(formula, Metadata);
        Assert.NotNull(result);

        var syntaxTree = CSharpSyntaxTree.ParseText(result.Source);
        var compilation = CSharpCompilation.Create(
            "DiagnosticCheck",
            new[] { syntaxTree },
            MetadataReferenceProvider.GetMetadataReferences(),
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        // Restrict to errors whose source span lies within the embedded user expression —
        // any errors elsewhere (e.g. ambiguous overloads in the synthetic stubs) are not what
        // this test verifies.
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
            $"Expected no diagnostic errors on '{innerExpression}' but got:\n{string.Join("\n", errors)}\n\nSource:\n{result.Source}");
    }
}
