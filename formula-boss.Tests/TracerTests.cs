using FormulaBoss.Runtime;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Unit tests for <see cref="Tracer" /> buffer mechanics: row cap + truncation,
///     column union across snapshots, cell keying, and clear-on-Begin.
/// </summary>
public class TracerTests : IDisposable
{
    public TracerTests()
    {
        Tracer.Reset();
    }

    public void Dispose()
    {
        Tracer.Reset();
    }

    [Fact]
    public void Begin_CreatesBuffer_SetsLastBuffer()
    {
        Tracer.Begin("foo", "Sheet1!A1");

        Assert.NotNull(Tracer.LastBuffer);
        Assert.Equal("foo", Tracer.LastBuffer!.Name);
        Assert.Equal("Sheet1!A1", Tracer.LastBuffer.CallerAddress);
        Assert.Empty(Tracer.LastBuffer.Rows);
    }

    [Fact]
    public void Begin_SecondCallWithSameCaller_ClearsBuffer()
    {
        Tracer.Begin("foo", "Sheet1!A1");
        Tracer.Set("x", 1);
        Tracer.Snapshot("iter", 0, null);
        Assert.Single(Tracer.LastBuffer!.Rows);

        Tracer.Begin("foo", "Sheet1!A1");
        Assert.Empty(Tracer.LastBuffer!.Rows);
        Assert.Empty(Tracer.LastBuffer.LocalNames);
    }

    [Fact]
    public void Begin_DifferentCallerAddress_NewBuffer()
    {
        Tracer.Begin("foo", "Sheet1!A1");
        var first = Tracer.LastBuffer;

        Tracer.Begin("foo", "Sheet1!B2");
        var second = Tracer.LastBuffer;

        Assert.NotSame(first, second);
        Assert.Equal("Sheet1!B2", second!.CallerAddress);
    }

    [Fact]
    public void Snapshot_CapturesLiveStateAndKindDepthBranch()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("x", 10);
        Tracer.Set("y", "hello");
        Tracer.Snapshot("iter", 1, "if (x>0)");

        var row = Tracer.LastBuffer!.Rows[0];
        Assert.Equal("iter", row["kind"]);
        Assert.Equal(1, row["depth"]);
        Assert.Equal("if (x>0)", row["branch"]);
        Assert.Equal(10, row["x"]);
        Assert.Equal("hello", row["y"]);
    }

    [Fact]
    public void Snapshot_CopiesLiveState_LaterSetDoesNotMutateRow()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("x", 1);
        Tracer.Snapshot("iter", 0, null);
        Tracer.Set("x", 2);

        Assert.Equal(1, Tracer.LastBuffer!.Rows[0]["x"]);
    }

    [Fact]
    public void Return_SetsReturnColumnInNextSnapshot()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("x", 42);
        Tracer.Return(99);
        Tracer.Snapshot("return", 0, null);

        var row = Tracer.LastBuffer!.Rows[0];
        Assert.Equal(99, row["return"]);
        Assert.Equal(42, row["x"]);
        Assert.Equal("return", row["kind"]);
    }

    [Fact]
    public void RowCap_TruncatesAt1000_AppendsWarningRow()
    {
        Tracer.Begin("foo", "A1");
        for (var i = 0; i < Tracer.MaxRows + 50; i++)
        {
            Tracer.Set("i", i);
            Tracer.Snapshot("iter", 0, null);
        }

        var rows = Tracer.LastBuffer!.Rows;
        Assert.Equal(Tracer.MaxRows + 1, rows.Count);
        var last = rows[rows.Count - 1];
        Assert.Equal("truncated", last["kind"]);
        Assert.Contains("1000", (string)last["branch"]!);
    }

    [Fact]
    public void TruncateWarn_AppendsSingleWarningRow_AndIsIdempotent()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("x", 1);
        Tracer.Snapshot("iter", 0, null);

        Tracer.TruncateWarn();
        Tracer.TruncateWarn();

        var rows = Tracer.LastBuffer!.Rows;
        Assert.Equal(2, rows.Count);
        Assert.Equal("truncated", rows[1]["kind"]);
    }

    [Fact]
    public void ToObjectArray_HeaderRowAndColumnUnion()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("a", 1);
        Tracer.Snapshot("entry", 0, null);
        Tracer.Set("b", 2);
        Tracer.Snapshot("iter", 0, null);
        Tracer.Set("c", 3);
        Tracer.Return(99);
        Tracer.Snapshot("return", 0, null);

        var arr = Tracer.LastBuffer!.ToObjectArray();

        Assert.Equal(4, arr.GetLength(0)); // header + 3 rows
        // Header: kind, depth, branch, a, b, c, return
        Assert.Equal("kind", arr[0, 0]);
        Assert.Equal("depth", arr[0, 1]);
        Assert.Equal("branch", arr[0, 2]);
        Assert.Equal("a", arr[0, 3]);
        Assert.Equal("b", arr[0, 4]);
        Assert.Equal("c", arr[0, 5]);
        Assert.Equal("return", arr[0, 6]);
    }

    [Fact]
    public void ToObjectArray_MissingCells_RenderAsEmptyString()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("a", 1);
        Tracer.Snapshot("entry", 0, null);
        Tracer.Set("b", 2);
        Tracer.Snapshot("iter", 0, null);

        var arr = Tracer.LastBuffer!.ToObjectArray();

        // Row 1 (entry) has a=1 but no b yet.
        // Columns: kind, depth, branch, a, b
        Assert.Equal(1, arr[1, 3]);
        Assert.Equal(string.Empty, arr[1, 4]);
        Assert.Equal(1, arr[2, 3]);
        Assert.Equal(2, arr[2, 4]);
    }

    [Fact]
    public void ToObjectArray_NoReturnColumn_WhenReturnNeverSet()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Set("a", 1);
        Tracer.Snapshot("iter", 0, null);

        var arr = Tracer.LastBuffer!.ToObjectArray();

        // Header: kind, depth, branch, a
        Assert.Equal(4, arr.GetLength(1));
        Assert.Equal("a", arr[0, 3]);
    }

    [Fact]
    public void ToObjectArray_NullBranch_RendersAsEmptyString()
    {
        Tracer.Begin("foo", "A1");
        Tracer.Snapshot("entry", 0, null);

        var arr = Tracer.LastBuffer!.ToObjectArray();
        Assert.Equal(string.Empty, arr[1, 2]);
    }

    [Fact]
    public void Set_WithoutBegin_DoesNotThrow()
    {
        Tracer.Reset();
        Tracer.Set("x", 1);
        Tracer.Snapshot("iter", 0, null);
        Tracer.Return(1);
        Tracer.TruncateWarn();
        Assert.Null(Tracer.LastBuffer);
    }
}
