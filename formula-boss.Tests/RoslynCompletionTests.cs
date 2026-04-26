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
    public async Task TableDot_ShowsColumnNames_AlongsideMembers()
    {
        var formula = "=`Sales.`";
        var textUp = "=`Sales.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        // Should show regular ExcelTable members
        Assert.Contains("Rows", texts);
        // Should also show column names
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);
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

    [Fact]
    public async Task RowAfterFirst_ShowsColumnCompletionsAsBracket()
    {
        var formula = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).`)";
        var textUp = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);

        // Column-name items insert as bracket syntax (ColumnCompletionData);
        // LINQ/ExcelArray members are regular CompletionData and must coexist.
        var columnItems = items.Where(i => i is ColumnCompletionData).Select(i => i.Text).ToList();
        Assert.Contains("Date", columnItems);
        Assert.Contains("Amount", columnItems);
        Assert.Contains("Region", columnItems);
    }

    [Fact]
    public async Task RowAfterFirst_AlsoShowsLinqAndExcelArrayMembers()
    {
        // Issue #325: a Row obtained via .First()/.Single() should expose its LINQ surface
        // (Skip, Where, FirstOrDefault, Map, ...) alongside the column headers.
        var formula = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).`)";
        var textUp = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Skip", texts);
        Assert.Contains("Where", texts);
        Assert.Contains("FirstOrDefault", texts);
        Assert.Contains("Map", texts);

        // LINQ members must be plain CompletionData so they insert as `.Skip(...)`,
        // not as bracket syntax.
        Assert.All(items.Where(i => i.Text is "Skip" or "Where" or "FirstOrDefault" or "Map"),
            item => Assert.IsNotType<ColumnCompletionData>(item));
    }

    [Fact]
    public async Task RowAfterFirstOrDefault_ShowsColumnCompletions()
    {
        var formula = "=LET(t, Sales, `t.Rows.FirstOrDefault(r => r[\"Region\"] == \"West\").`)";
        var textUp = "=LET(t, Sales, `t.Rows.FirstOrDefault(r => r[\"Region\"] == \"West\").";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);

        var columnItems = items.Where(i => i is ColumnCompletionData).Select(i => i.Text).ToList();
        Assert.Contains("Date", columnItems);
        Assert.Contains("Amount", columnItems);
        Assert.Contains("Region", columnItems);
    }

    [Fact]
    public async Task RowLambdaParam_ShowsColumnsAndLinqMembers()
    {
        // Lambda parameter on a Row context (.Rows.Where(r => r.)) should also expose
        // the LINQ surface alongside columns.
        var formula = "=LET(t, Sales, `t.Rows.Where(r => r.`)";
        var textUp = "=LET(t, Sales, `t.Rows.Where(r => r.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Skip", texts);
        Assert.Contains("FirstOrDefault", texts);
    }

    [Fact]
    public async Task Row_ColumnNamesNotDuplicated()
    {
        // The synthetic Row class emits column-name properties for Roslyn; the bracket-
        // syntax versions come from BuildRowCompletions. The merge must not list a
        // column twice.
        var formula = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).`)";
        var textUp = "=LET(t, Sales, `t.Rows.First(r => r[\"Amount\"] > 100).";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.Single(items, i => i.Text == "Date");
        Assert.Single(items, i => i.Text == "Amount");
        Assert.Single(items, i => i.Text == "Region");
    }

    [Fact]
    public async Task NonRowType_NotAffected()
    {
        // String.Length is not a Row — should still return normal Roslyn completions
        var formula = "=`\"hello\".`";
        var textUp = "=`\"hello\".";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Length", texts);

        // These should be regular CompletionData, not ColumnCompletionData
        Assert.All(items, item => Assert.IsNotType<ColumnCompletionData>(item));
    }

    [Fact]
    public async Task GroupBySelect_ShowsRowGroupMembers_NotColumnNames()
    {
        var formula = "=LET(t, Sales, `t.Rows.GroupBy(r => r[\"Region\"]).Select(g => g.`)";
        var textUp = "=LET(t, Sales, `t.Rows.GroupBy(r => r[\"Region\"]).Select(g => g.";

        var (items, isBracket) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.False(isBracket);
        var texts = items.Select(i => i.Text).ToList();

        // Should show RowGroup members
        Assert.Contains("Key", texts);
        Assert.Contains("Count", texts);
        Assert.Contains("Where", texts);
        Assert.Contains("Select", texts);
        Assert.Contains("ToRange", texts);

        // Should NOT show column names (those belong to Row, not RowGroup)
        Assert.DoesNotContain("Date", texts);
        Assert.DoesNotContain("Amount", texts);
        Assert.DoesNotContain("Region", texts);

        // Should be regular CompletionData, not ColumnCompletionData
        Assert.All(items, item => Assert.IsNotType<ColumnCompletionData>(item));
    }

    [Fact]
    public async Task GroupByWhere_ShowsRowGroupMembers()
    {
        var formula = "=LET(t, Sales, `t.Rows.GroupBy(r => r[\"Region\"]).Where(g => g.`)";
        var textUp = "=LET(t, Sales, `t.Rows.GroupBy(r => r[\"Region\"]).Where(g => g.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Key", texts);
        Assert.Contains("Count", texts);
        Assert.DoesNotContain("Date", texts);
    }

    [Fact]
    public async Task GroupByLambdaParam_InGroupBy_StillShowsColumns()
    {
        // The lambda parameter inside GroupBy itself is a Row — should show column completions
        var formula = "=LET(t, Sales, `t.Rows.GroupBy(r => r.`)";
        var textUp = "=LET(t, Sales, `t.Rows.GroupBy(r => r.";

        var (items, _) = await _provider.GetCompletionsAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        var texts = items.Select(i => i.Text).ToList();
        Assert.Contains("Date", texts);
        Assert.Contains("Amount", texts);
        Assert.Contains("Region", texts);
    }
}
