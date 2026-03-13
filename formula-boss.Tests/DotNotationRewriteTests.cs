using FormulaBoss.Transpilation;

using Xunit;

namespace FormulaBoss.Tests;

public class ColumnMapperTests
{
    [Theory]
    [InlineData("Country", "Country")]
    [InlineData("Population 2025", "Population2025")]
    [InlineData("GDP (USD)", "GDPUSD")]
    [InlineData("First Name", "FirstName")]
    [InlineData("col_name", "col_name")]
    [InlineData("123start", "_123start")]
    [InlineData("already_valid", "already_valid")]
    public void Sanitise_ProducesValidCSharpIdentifier(string input, string expected)
    {
        Assert.Equal(expected, ColumnMapper.Sanitise(input));
    }

    [Fact]
    public void Sanitise_EmptyOrSpecialOnly_ReturnsEmpty()
    {
        Assert.Equal("", ColumnMapper.Sanitise(""));
        Assert.Equal("", ColumnMapper.Sanitise("!@#$%"));
    }

    [Fact]
    public void BuildMapping_SimpleCases()
    {
        var headers = new[] { "Country", "Population 2025", "GDP" };
        var mapping = ColumnMapper.BuildMapping(headers);

        Assert.Equal(3, mapping.Count);
        Assert.Equal("Country", mapping["Country"]);
        Assert.Equal("Population 2025", mapping["Population2025"]);
        Assert.Equal("GDP", mapping["GDP"]);
    }

    [Fact]
    public void BuildMapping_ConflictDetection_ExcludesBoth()
    {
        // "Foo Bar" and "FooBar" both sanitise to "FooBar"
        var headers = new[] { "Foo Bar", "FooBar", "Safe Column" };
        var mapping = ColumnMapper.BuildMapping(headers);

        Assert.False(mapping.ContainsKey("FooBar"));
        Assert.Equal("Safe Column", mapping["SafeColumn"]);
        Assert.Single(mapping);
    }

    [Fact]
    public void BuildMapping_AlreadyValidIdentifiers_MapToThemselves()
    {
        var headers = new[] { "Price", "Quantity", "Total" };
        var mapping = ColumnMapper.BuildMapping(headers);

        Assert.Equal("Price", mapping["Price"]);
        Assert.Equal("Quantity", mapping["Quantity"]);
        Assert.Equal("Total", mapping["Total"]);
    }

    [Fact]
    public void BuildMapping_EmptyHeaders_ReturnsEmptyMapping()
    {
        var mapping = ColumnMapper.BuildMapping(Array.Empty<string>());
        Assert.Empty(mapping);
    }

    [Fact]
    public void BuildMapping_SpecialCharsOnly_ExcludedFromMapping()
    {
        var headers = new[] { "!@#", "Valid" };
        var mapping = ColumnMapper.BuildMapping(headers);

        Assert.Single(mapping);
        Assert.Equal("Valid", mapping["Valid"]);
    }
}

public class DotNotationRewriterTests
{
    [Fact]
    public void Rewrite_SimpleDotAccess_ConvertsToBracket()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Population 2025" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Population2025 > 1000)",
            mapping);

        Assert.Contains("r[\"Population 2025\"]", result);
        Assert.DoesNotContain("r.Population2025", result);
    }

    [Fact]
    public void Rewrite_MultipleDotAccesses_ConvertsAll()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "First Name", "Last Name" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.FirstName == \"Tim\" && r.LastName == \"Smith\")",
            mapping);

        Assert.Contains("r[\"First Name\"]", result);
        Assert.Contains("r[\"Last Name\"]", result);
    }

    [Fact]
    public void Rewrite_MixedDotAndBracket_OnlyRewritesDot()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Population 2025", "Country" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Population2025 > 1000 && r[\"Country\"] == \"US\")",
            mapping);

        Assert.Contains("r[\"Population 2025\"]", result);
        // Existing bracket access preserved
        Assert.Contains("r[\"Country\"]", result);
    }

    [Fact]
    public void Rewrite_AlreadyValidIdentifier_NoOpMapping()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Price", "Quantity" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Price > 5)",
            mapping);

        // "Price" sanitises to "Price", so r.Price → r["Price"]
        Assert.Contains("r[\"Price\"]", result);
    }

    [Fact]
    public void Rewrite_NoMapping_ReturnsUnchanged()
    {
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Price > 5)",
            new Dictionary<string, string>());

        Assert.Contains("r.Price", result);
    }

    [Fact]
    public void Rewrite_NonRowMemberAccess_NotRewritten()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Rows", "Count" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Count > 0)",
            mapping);

        // tbl.Rows should NOT be rewritten (tbl is not a lambda param)
        Assert.Contains("tbl.Rows", result);
        // r.Count SHOULD be rewritten (r is a lambda param and Count is in mapping)
        Assert.Contains("r[\"Count\"]", result);
    }

    [Fact]
    public void Rewrite_TableParameter_DotNotation_ConvertsToBracket()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Price", "Quantity" });
        var tableParamNames = new HashSet<string> { "tbl" };
        var result = DotNotationRewriter.Rewrite(
            "tbl.Price",
            mapping,
            tableParamNames);

        Assert.Contains("tbl[\"Price\"]", result);
        Assert.DoesNotContain("tbl.Price", result);
    }

    [Fact]
    public void Rewrite_TableParameter_MergesWithLambdaParams()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Age", "Name" });
        var tableParamNames = new HashSet<string> { "tbl" };
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Age > tbl.Name)",
            mapping,
            tableParamNames);

        // Both lambda param and table param should be rewritten
        Assert.Contains("r[\"Age\"]", result);
        Assert.Contains("tbl[\"Name\"]", result);
    }

    [Fact]
    public void Rewrite_ConflictingColumns_NotRewritten()
    {
        // "Foo Bar" and "FooBar" both → "FooBar", so conflict → excluded
        var mapping = ColumnMapper.BuildMapping(new[] { "Foo Bar", "FooBar", "Safe" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.FooBar > 0 && r.Safe == true)",
            mapping);

        // FooBar is conflicted, should NOT be rewritten
        Assert.Contains("r.FooBar", result);
        // Safe is not conflicted, SHOULD be rewritten
        Assert.Contains("r[\"Safe\"]", result);
    }

    [Fact]
    public void Rewrite_NestedLambdas_RewritesBoth()
    {
        var mapping = ColumnMapper.BuildMapping(new[] { "Population 2025", "Continent" });
        var result = DotNotationRewriter.Rewrite(
            "tbl.Rows.Where(r => r.Population2025 > 1000 && conts.Any(c => c.Continent == \"EU\"))",
            mapping);

        Assert.Contains("r[\"Population 2025\"]", result);
        Assert.Contains("c[\"Continent\"]", result);
    }
}

public class CodeEmitterDotNotationTests
{
    [Fact]
    public void Emit_WithHeaders_RewritesDotNotation()
    {
        var detection = new DetectionResult(
            Parameters: new List<string> { "tbl" },
            RequiresObjectModel: false,
            HeaderVariables: new HashSet<string> { "tbl" },
            NormalizedExpression: "tbl.Rows.Where(r => r.Population2025 > 1000).ToResult()",
            RangeRefMap: new Dictionary<string, string>());

        var headers = new Dictionary<string, string[]>
        {
            ["tbl"] = new[] { "Country", "Population 2025" }
        };

        var emitter = new CodeEmitter();
        var result = emitter.Emit(detection, "original", "test", headers);

        Assert.Contains("r[\"Population 2025\"]", result.SourceCode);
        Assert.DoesNotContain("r.Population2025", result.SourceCode);
    }

    [Fact]
    public void Emit_WithoutHeaders_NoRewrite()
    {
        var detection = new DetectionResult(
            Parameters: new List<string> { "tbl" },
            RequiresObjectModel: false,
            HeaderVariables: new HashSet<string> { "tbl" },
            NormalizedExpression: "tbl.Rows.Where(r => r.Price > 5).ToResult()",
            RangeRefMap: new Dictionary<string, string>());

        var emitter = new CodeEmitter();
        var result = emitter.Emit(detection, "original", "test");

        Assert.Contains("r.Price", result.SourceCode);
    }
}
