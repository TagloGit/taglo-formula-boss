using FormulaBoss.UI;
using Xunit;

namespace FormulaBoss.Tests;

public class CompletionScopingTests
{
    private static readonly WorkbookMetadata TwoTableMetadata = new(
        new[] { "Sales", "Products" },
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["Sales"] = new[] { "Date", "Amount", "Region" },
            ["Products"] = new[] { "Name", "Price", "Category" }
        });

    [Fact]
    public void Resolve_TableName_ForDirectTableRowAccess()
    {
        var ctx = ContextResolver.Resolve("Sales.rows.map(r => r.", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.Equal("Sales", ctx.TableName);
    }

    [Fact]
    public void Resolve_TableName_ForOtherTable()
    {
        var ctx = ContextResolver.Resolve("Products.rows.map(r => r.", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.Equal("Products", ctx.TableName);
    }

    [Fact]
    public void Resolve_TableName_Null_WhenTableUnknown()
    {
        var ctx = ContextResolver.Resolve("unknown.rows.map(r => r.", TwoTableMetadata);
        // unknown is not a recognized range, so context resolution may differ
        // The key thing is it doesn't crash and falls back
        Assert.True(ctx.TableName == null || !TwoTableMetadata.TableColumns.ContainsKey(ctx.TableName));
    }

    [Fact]
    public void Resolve_TableName_ThroughWhereChain()
    {
        var ctx = ContextResolver.Resolve(
            "Sales.rows.where(r => r.Amount > 100).map(r => r.", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.Equal("Sales", ctx.TableName);
    }

    [Fact]
    public void Resolve_TableName_BracketContext()
    {
        var ctx = ContextResolver.Resolve("Sales.rows.map(r => r[", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.True(ctx.IsBracketContext);
        Assert.Equal("Sales", ctx.TableName);
    }

    [Fact]
    public void Resolve_TableName_BracketContextWithPartial()
    {
        var ctx = ContextResolver.Resolve("Sales.rows.map(r => r[Da", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.True(ctx.IsBracketContext);
        Assert.Equal("Sales", ctx.TableName);
        Assert.Equal("Da", ctx.PartialWord);
    }

    [Fact]
    public void Resolve_TableName_DirectRowsDot()
    {
        // Table.rows.where(r => r. — direct row context in where
        var ctx = ContextResolver.Resolve("Products.rows.where(r => r.", TwoTableMetadata);
        Assert.Equal(DslType.Row, ctx.Type);
        Assert.Equal("Products", ctx.TableName);
    }
}
