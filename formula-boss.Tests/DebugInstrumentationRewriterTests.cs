using FormulaBoss.Transpilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

using Xunit;

namespace FormulaBoss.Tests;

public class DebugInstrumentationRewriterTests
{
    private static string Rewrite(string blockText, string name = "TEST", string callerAddr = "\"A1\"")
    {
        var block = (BlockSyntax)SyntaxFactory.ParseStatement(blockText);
        var rewritten = DebugInstrumentationRewriter.Instrument(block, name, callerAddr);
        return rewritten.NormalizeWhitespace().ToFullString();
    }

    [Fact]
    public void Instrument_PrependsBeginAndEntrySnapshot()
    {
        var code = Rewrite("{ return 1; }");

        Assert.Contains("Tracer.Begin(\"TEST\", \"A1\")", code);
        Assert.Contains("Tracer.Snapshot(\"entry\", 0, null)", code);
    }

    [Fact]
    public void Instrument_EmitsSetAfterLocalDeclaration()
    {
        var code = Rewrite("{ var x = 1; return x; }");

        Assert.Contains("Tracer.Set(\"x\", x)", code);
    }

    [Fact]
    public void Instrument_EmitsSetForEachVariableInMultiDeclaration()
    {
        var code = Rewrite("{ var x = 1; var y = 2; return x + y; }");

        Assert.Contains("Tracer.Set(\"x\", x)", code);
        Assert.Contains("Tracer.Set(\"y\", y)", code);
    }

    [Fact]
    public void Instrument_EmitsSetAfterSimpleAssignment()
    {
        var code = Rewrite("{ var x = 1; x = 2; return x; }");

        // One after declaration, one after assignment.
        Assert.True(code.Split("Tracer.Set(\"x\", x)").Length - 1 >= 2);
    }

    [Fact]
    public void Instrument_ForeachBindsLoopVariableAndSnapshotsAtEnd()
    {
        var code = Rewrite("{ foreach (var r in rows) { var a = r; } return 0; }");

        Assert.Contains("Tracer.Set(\"r\", r)", code);
        Assert.Contains("Tracer.Set(\"a\", a)", code);
        Assert.Contains("Tracer.Snapshot(\"iter\", 1, null)", code);
    }

    [Fact]
    public void Instrument_NestedForeachTracksDepth()
    {
        var code = Rewrite(
            "{ foreach (var r in rows) { foreach (var c in cols) { var v = 1; } } return 0; }");

        Assert.Contains("Tracer.Snapshot(\"iter\", 1, null)", code);
        Assert.Contains("Tracer.Snapshot(\"iter\", 2, null)", code);
    }

    [Fact]
    public void Instrument_IfElse_TagsBranchLabels()
    {
        var code = Rewrite(
            "{ if (a > 0) { return 1; } else { return -1; } }");

        // Branch label derived from first statement of the arm.
        Assert.Contains("Tracer.Snapshot(\"return\", 0, \"return 1;\")", code);
        Assert.Contains("Tracer.Snapshot(\"return\", 0, \"return -1;\")", code);
    }

    [Fact]
    public void Instrument_IfWithoutElse_LabelsThenArm()
    {
        var code = Rewrite("{ if (x > 0) { return x; } return 0; }");

        Assert.Contains("Tracer.Snapshot(\"return\", 0, \"return x;\")", code);
        // The trailing return is outside the if — branch label is null.
        Assert.Contains("Tracer.Snapshot(\"return\", 0, null)", code);
    }

    [Fact]
    public void Instrument_ReturnInLoop_UsesLoopDepthAndCapturesValue()
    {
        var code = Rewrite("{ foreach (var r in rows) { return r; } return 0; }");

        Assert.Contains("var __fb_ret = r", code);
        Assert.Contains("Tracer.Return(__fb_ret)", code);
        Assert.Contains("Tracer.Snapshot(\"return\", 1, null)", code);
        Assert.Contains("return __fb_ret;", code);
    }

    [Fact]
    public void Instrument_EarlyReturn_CapturesValueBeforeReturning()
    {
        var code = Rewrite("{ var x = 5; if (x > 0) return x; return 0; }");

        Assert.Contains("var __fb_ret = x", code);
        Assert.Contains("Tracer.Return(__fb_ret)", code);
        Assert.Contains("return __fb_ret;", code);
    }

    [Fact]
    public void Instrument_ForLoop_TracksDeclaredVariableAndSnapshots()
    {
        var code = Rewrite("{ for (var i = 0; i < 10; i++) { var v = i; } return 0; }");

        Assert.Contains("Tracer.Set(\"i\", i)", code);
        Assert.Contains("Tracer.Set(\"v\", v)", code);
        Assert.Contains("Tracer.Snapshot(\"iter\", 1, null)", code);
    }

    [Fact]
    public void Instrument_WhileLoop_SnapshotsAtEnd()
    {
        var code = Rewrite("{ var i = 0; while (i < 10) { i = i + 1; } return i; }");

        Assert.Contains("Tracer.Snapshot(\"iter\", 1, null)", code);
    }

    [Fact]
    public void Instrument_BranchLabelTruncatedToMaxLength()
    {
        var code = Rewrite(
            "{ if (a) { var longVariableNameThatExceedsLimit = 1; return 1; } return 0; }");

        // The label should appear truncated to 20 chars — the full decl statement should NOT appear
        // in any Snapshot call.
        Assert.DoesNotContain(
            "Tracer.Snapshot(\"return\", 0, \"var longVariableNameThatExceedsLimit = 1;\"",
            code);
    }
}
