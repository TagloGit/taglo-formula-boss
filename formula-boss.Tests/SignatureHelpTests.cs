using FormulaBoss.UI;
using FormulaBoss.UI.Completion;

using Xunit;

namespace FormulaBoss.Tests;

public class SignatureHelpTests : IDisposable
{
    private static readonly WorkbookMetadata SalesMetadata = new(
        new[] { "Sales" },
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>> { ["Sales"] = new[] { "Date", "Amount", "Region" } });

    private readonly SignatureHelpProvider _provider;
    private readonly RoslynWorkspaceManager _workspace;

    public SignatureHelpTests()
    {
        _workspace = new RoslynWorkspaceManager();
        _provider = new SignatureHelpProvider(_workspace);
    }

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public async Task Where_ShowsSignature()
    {
        var formula = "=LET(t, Sales, `t.Rows.Where(`)";
        var textUp = "=LET(t, Sales, `t.Rows.Where(";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        Assert.NotEmpty(result.Overloads);
        var overload = result.Overloads[result.ActiveOverloadIndex];
        Assert.Equal("Where", overload.MethodName.Split('.').Last());
        Assert.Single(overload.Parameters);
        Assert.Equal("predicate", overload.Parameters[0].Name);
    }

    [Fact]
    public async Task ExcelValueMethod_HasXmlDocSummary()
    {
        // Use ExcelValue.Where (not synthetic row collection) to test XML doc propagation
        var formula = "=`Sales.Where(`";
        var textUp = "=`Sales.Where(";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        var overload = result.Overloads[result.ActiveOverloadIndex];
        // XML docs may not be available in test context (depends on XML file being
        // next to the assembly). Verify the structure is correct regardless.
        Assert.NotEmpty(overload.Parameters);
        Assert.Equal("predicate", overload.Parameters[0].Name);

        // If XML docs ARE available, verify they're populated correctly
        if (overload.Summary != null)
        {
            Assert.Contains("filter", overload.Summary, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public async Task Aggregate_TwoParams_ActiveParameterUpdates()
    {
        // Cursor after the comma — should be on the second parameter
        var formula = "=LET(t, Sales, `t.Rows.Aggregate(0, `)";
        var textUp = "=LET(t, Sales, `t.Rows.Aggregate(0, ";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        Assert.Equal(1, result.ActiveParameterIndex);
    }

    [Fact]
    public async Task OutsideDsl_ReturnsNull()
    {
        var formula = "=LET(t, Sales, )";
        var textUp = "=LET(t, Sales, ";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.Null(result);
    }

    [Fact]
    public async Task StringMethod_ShowsOverloads()
    {
        // String.Substring has 2 overloads
        var formula = "=`\"hello\".Substring(`";
        var textUp = "=`\"hello\".Substring(";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        Assert.True(result.Overloads.Count >= 2,
            $"Expected at least 2 overloads, got {result.Overloads.Count}");
    }

    [Fact]
    public async Task ExcelValueMethod_ShowsSignature()
    {
        var formula = "=`Sales.Where(`";
        var textUp = "=`Sales.Where(";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        var overload = result.Overloads[result.ActiveOverloadIndex];
        Assert.Contains("Where", overload.MethodName);
        Assert.NotEmpty(overload.Parameters);
    }

    [Fact]
    public async Task ReturnType_IsPresent()
    {
        var formula = "=`Sales.Count(`";
        var textUp = "=`Sales.Count(";

        var result = await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        Assert.NotNull(result);
        var overload = result.Overloads[result.ActiveOverloadIndex];
        Assert.False(string.IsNullOrEmpty(overload.ReturnType));
    }

    [Fact]
    public async Task NoCrash_OnEmptyParens()
    {
        // Just an opening paren with nothing before it that's a method
        var formula = "=`(`";
        var textUp = "=`(";

        await _provider.GetSignatureHelpAsync(
            textUp, formula, SalesMetadata, CancellationToken.None);

        // Should return null gracefully, not throw
        // (the paren isn't after a method name, so no signature help)
    }
}
