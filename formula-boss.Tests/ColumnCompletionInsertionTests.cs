using FormulaBoss.UI;

using Xunit;

namespace FormulaBoss.Tests;

public class ColumnCompletionInsertionTests
{
    private static string Apply(string documentText, string columnName, bool isBracket, int segmentOffset, int segmentLength)
    {
        var (newText, replaceOffset, replaceLength) =
            ColumnCompletionData.ComputeInsertion(columnName, isBracket, documentText, segmentOffset, segmentLength);
        return documentText[..replaceOffset] + newText + documentText[(replaceOffset + replaceLength)..];
    }

    // --- Bracket context ---

    [Fact]
    public void Bracket_WithAutoClosedBracket_NoDoubleBracket()
    {
        // User typed r[, auto-closer added ], cursor is between: r[|]
        // Segment is empty (offset=2, length=0), ] is at position 2
        var result = Apply("r[]", "Amount", isBracket: true, segmentOffset: 2, segmentLength: 0);
        Assert.Equal("r[\"Amount\"]", result);
    }

    [Fact]
    public void Bracket_WithPartialTyping_AndAutoClosedBracket()
    {
        // User typed r[Am, auto-closer added ]: r[Am|]
        // Segment covers "Am" (offset=2, length=2), ] at position 4
        var result = Apply("r[Am]", "Amount", isBracket: true, segmentOffset: 2, segmentLength: 2);
        Assert.Equal("r[\"Amount\"]", result);
    }

    [Fact]
    public void Bracket_WithoutAutoClosedBracket()
    {
        // No auto-closer (e.g. manually deleted it): r[|
        var result = Apply("r[", "Amount", isBracket: true, segmentOffset: 2, segmentLength: 0);
        Assert.Equal("r[\"Amount\"]", result);
    }

    [Fact]
    public void Bracket_ColumnNameWithSpaces()
    {
        var result = Apply("r[]", "First Name", isBracket: true, segmentOffset: 2, segmentLength: 0);
        Assert.Equal("r[\"First Name\"]", result);
    }

    // --- Dot context ---

    [Fact]
    public void Dot_RewritesToBracketSyntax()
    {
        // User typed r. then completion fires with empty segment after dot
        var result = Apply("r.", "Amount", isBracket: false, segmentOffset: 2, segmentLength: 0);
        Assert.Equal("r[\"Amount\"]", result);
    }

    [Fact]
    public void Dot_WithPartialTyping()
    {
        // User typed r.Am, segment covers "Am" (offset=2, length=2)
        var result = Apply("r.Am", "Amount", isBracket: false, segmentOffset: 2, segmentLength: 2);
        Assert.Equal("r[\"Amount\"]", result);
    }

    [Fact]
    public void Dot_ColumnNameWithSpaces()
    {
        var result = Apply("r.", "First Name", isBracket: false, segmentOffset: 2, segmentLength: 0);
        Assert.Equal("r[\"First Name\"]", result);
    }

    [Fact]
    public void Dot_PreservesTextBeforeAndAfter()
    {
        // Expression with surrounding context
        var result = Apply("tbl.rows.where(r => r. > 5)", "Amount", isBracket: false, segmentOffset: 22, segmentLength: 0);
        Assert.Equal("tbl.rows.where(r => r[\"Amount\"] > 5)", result);
    }
}
