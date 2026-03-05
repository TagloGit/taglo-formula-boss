using FormulaBoss.UI;
using FormulaBoss.UI.Completion;

using Xunit;

namespace FormulaBoss.Tests;

public class RoslynCompletionTests : IDisposable
{
    private readonly RoslynWorkspaceManager _workspace;
    private readonly RoslynCompletionProvider _provider;

    private static readonly WorkbookMetadata SalesMetadata = new(
        new[] { "Sales" },
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Sales"] = new[] { "Date", "Amount", "Region" }
        });

    public RoslynCompletionTests()
    {
        _workspace = new RoslynWorkspaceManager();
        _provider = new RoslynCompletionProvider(_workspace);
    }

    public void Dispose()
    {
        _workspace.Dispose();
    }

    [Fact]
    public async Task TableDot_ShowsAccessorsAndMethods()
    {
        var formula = "=`Sales.`";
        var textUp = "=`Sales.";

        var (items, isBracket) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.False(isBracket);
        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Rows", texts);
        Assert.Contains("Cells", texts);
        Assert.Contains("Where", texts);
    }

    [Fact]
    public async Task RowDot_ShowsColumnProperties()
    {
        var formula = "=LET(t, Sales, `t.Rows.Where(r => r.`)";
        var textUp = "=LET(t, Sales, `t.Rows.Where(r => r.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);
    }

    [Fact]
    public async Task BracketContext_UsesFallback()
    {
        var formula = "=`Sales.rows.map(r => r[`";
        var textUp = "=`Sales.rows.map(r => r[";

        var (items, isBracket) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.True(isBracket);
        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);
    }

    [Fact]
    public async Task NonDslContext_ReturnsTableNames()
    {
        var formula = "=LET(t, S";
        var textUp = "=LET(t, S";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Sales", texts);
    }

    [Fact]
    public async Task StringMethods_Available()
    {
        var formula = "=`\"hello\".`";
        var textUp = "=`\"hello\".";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Length", texts);
        Assert.Contains("Substring", texts);
    }

    [Fact]
    public async Task RegexDot_ShowsStaticMethods()
    {
        var formula = "=`Regex.`";
        var textUp = "=`Regex.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Match", texts);
        Assert.Contains("Replace", texts);
        Assert.Contains("IsMatch", texts);
    }

    [Fact]
    public async Task FiltersInternalTypes()
    {
        var formula = "=`Sales.`";
        var textUp = "=`Sales.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.DoesNotContain("__SalesRow", texts);
        Assert.DoesNotContain("__SalesTable", texts);
        Assert.DoesNotContain("__Ctx", texts);
    }

    [Fact]
    public async Task FiltersObjectMethods()
    {
        var formula = "=`Sales.`";
        var textUp = "=`Sales.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.DoesNotContain("GetHashCode", texts);
        Assert.DoesNotContain("GetType", texts);
    }

    [Fact]
    public async Task StatementBlock_ShowsCompletions()
    {
        var formula = "=LET(myFunc, `{ var s = new StringBuilder(); s. }`)";
        var textUp = "=LET(myFunc, `{ var s = new StringBuilder(); s.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Length", texts);
        Assert.Contains("Append", texts);
        Assert.Contains("AppendLine", texts);
    }

    [Fact]
    public async Task StatementBlock_Multiline_ShowsCompletions()
    {
        var formula = "=LET(myFunc, `{\n  var s = new StringBuilder();\n  s.\n}`)";
        var textUp = "=LET(myFunc, `{\n  var s = new StringBuilder();\n  s.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Length", texts);
        Assert.Contains("Append", texts);
        Assert.Contains("AppendLine", texts);
    }

    [Fact]
    public async Task NonDslContext_ShowsLetVariables_WhenFormulaContainsBackticks()
    {
        var formula = "=LET(input, A1, myFormula, `Sales.Rows.Where(r => r.Amount > 100)`, )";
        var textUp = "=LET(input, A1, myFormula, `Sales.Rows.Where(r => r.Amount > 100)`, ";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("input", texts);
        Assert.Contains("myFormula", texts);
    }
}
