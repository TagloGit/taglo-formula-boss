using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

using FormulaBoss.UI.Completion;

using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;

namespace FormulaBoss.UI;

/// <summary>
///     Highlights Roslyn compile errors in backtick expressions with red squiggly underlines.
///     Shows tooltip messages on mouse hover over error regions.
/// </summary>
internal sealed class ErrorHighlighter : IBackgroundRenderer
{
    private readonly TextEditor _editor;
    private readonly DispatcherTimer _debounceTimer;
    private readonly Pen _squigglePen;
    private readonly ToolTip _tooltip;
    private readonly Func<RoslynWorkspaceManager?> _getWorkspace;
    private readonly Func<WorkbookMetadata?> _getMetadata;

    private CancellationTokenSource? _diagnosticCts;
    private List<ErrorMarker> _markers = [];

    public ErrorHighlighter(
        TextEditor editor,
        Func<RoslynWorkspaceManager?> getWorkspace,
        Func<WorkbookMetadata?> getMetadata)
    {
        _editor = editor;
        _getWorkspace = getWorkspace;
        _getMetadata = getMetadata;

        var red = Color.FromRgb(0xFF, 0x30, 0x30);
        _squigglePen = new Pen(new SolidColorBrush(red), 1.2) { DashStyle = DashStyles.Dot };
        _squigglePen.Freeze();

        _tooltip = new ToolTip { Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse };

        _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _debounceTimer.Tick += OnDebounceTimerTick;

        editor.TextArea.TextView.BackgroundRenderers.Add(this);
        editor.TextChanged += (_, _) =>
        {
            _debounceTimer.Stop();
            _debounceTimer.Start();
        };

        editor.TextArea.TextView.MouseMove += OnMouseMove;
        editor.TextArea.TextView.MouseLeave += (_, _) => _tooltip.IsOpen = false;
    }

    public KnownLayer Layer => KnownLayer.Selection;

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (_markers.Count == 0)
        {
            return;
        }

        foreach (var marker in _markers)
        {
            // Clamp to document bounds
            var start = Math.Min(marker.StartOffset, _editor.Document.TextLength);
            var length = Math.Min(marker.Length, _editor.Document.TextLength - start);
            if (length <= 0)
            {
                continue;
            }

            var segment = new TextSegment { StartOffset = start, Length = length };
            foreach (var rect in BackgroundGeometryBuilder.GetRectsForSegment(textView, segment))
            {
                DrawSquiggly(drawingContext, rect);
            }
        }
    }

    private void DrawSquiggly(DrawingContext dc, Rect rect)
    {
        var start = new Point(rect.Left, rect.Bottom);
        var end = new Point(rect.Right, rect.Bottom);

        const double waveHeight = 2.0;
        const double waveLength = 4.0;

        var geometry = new StreamGeometry();
        using (var ctx = geometry.Open())
        {
            ctx.BeginFigure(start, false, false);
            var x = start.X;
            var up = true;
            while (x < end.X)
            {
                var nextX = Math.Min(x + waveLength, end.X);
                var y = up ? start.Y - waveHeight : start.Y;
                ctx.LineTo(new Point(nextX, y), true, false);
                x = nextX;
                up = !up;
            }
        }

        geometry.Freeze();
        dc.DrawGeometry(null, _squigglePen, geometry);
    }

    private async void OnDebounceTimerTick(object? sender, EventArgs e)
    {
        _debounceTimer.Stop();
        _diagnosticCts?.Cancel();
        _diagnosticCts = new CancellationTokenSource();
        var ct = _diagnosticCts.Token;

        try
        {
            await UpdateErrorsAsync(ct);
        }
        catch (OperationCanceledException)
        {
            // Expected when user types quickly
        }
    }

    private async Task UpdateErrorsAsync(CancellationToken ct)
    {
        var workspace = _getWorkspace();
        if (workspace == null)
        {
            return;
        }

        var text = _editor.Text;
        var metadata = _getMetadata();

        var buildResult = SyntheticDocumentBuilder.BuildForDiagnostics(text, metadata);
        if (buildResult == null)
        {
            _markers = [];
            _editor.TextArea.TextView.InvalidateLayer(Layer);
            return;
        }

        ct.ThrowIfCancellationRequested();

        var diagnostics = await workspace.GetDiagnosticsAsync(buildResult.Source, ct);

        ct.ThrowIfCancellationRequested();

        var newMarkers = new List<ErrorMarker>();
        var exprStart = buildResult.ExpressionStartInSynthetic;
        var exprEnd = exprStart + buildResult.ExpressionLength;

        foreach (var diagnostic in diagnostics)
        {
            var span = diagnostic.Location.SourceSpan;

            // Only show diagnostics within the expression region
            if (span.Start < exprStart || span.Start >= exprEnd)
            {
                continue;
            }

            // Filter trailing diagnostics to reduce noise while typing
            if (span.Start >= exprEnd - 2)
            {
                continue;
            }

            var editorOffset = (span.Start - exprStart) + buildResult.ExpressionStartInEditor;
            var length = Math.Min(span.Length, buildResult.ExpressionLength - (span.Start - exprStart));
            if (length <= 0)
            {
                length = 1;
            }

            newMarkers.Add(new ErrorMarker(editorOffset, length, diagnostic.GetMessage()));
        }

        _markers = newMarkers;
        _editor.TextArea.TextView.InvalidateLayer(Layer);
    }

    private void OnMouseMove(object sender, MouseEventArgs e)
    {
        if (_markers.Count == 0)
        {
            _tooltip.IsOpen = false;
            return;
        }

        var pos = _editor.TextArea.TextView.GetPosition(e.GetPosition(_editor.TextArea.TextView));
        if (pos == null)
        {
            _tooltip.IsOpen = false;
            return;
        }

        var offset = _editor.Document.GetOffset(pos.Value.Location);
        var hit = _markers.Find(m => offset >= m.StartOffset && offset < m.StartOffset + m.Length);

        if (hit != null)
        {
            _tooltip.Content = hit.Message;
            _tooltip.IsOpen = true;
        }
        else
        {
            _tooltip.IsOpen = false;
        }
    }

    private sealed record ErrorMarker(int StartOffset, int Length, string Message);
}
