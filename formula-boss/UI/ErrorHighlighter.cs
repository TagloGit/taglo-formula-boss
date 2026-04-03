using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

using FormulaBoss.Analysis;
using FormulaBoss.Interception;
using FormulaBoss.UI.Completion;

using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;

using Microsoft.CodeAnalysis;

namespace FormulaBoss.UI;

/// <summary>
///     Highlights Roslyn compile errors in backtick expressions with red squiggly underlines.
///     Shows error tooltip messages on hover over error regions and type tooltips on hover
///     over identifiers, var keywords, and lambda parameters.
/// </summary>
internal sealed class ErrorHighlighter : IBackgroundRenderer
{
    private readonly DispatcherTimer _debounceTimer;
    private readonly TextEditor _editor;
    private readonly Func<WorkbookMetadata?> _getMetadata;
    private readonly Func<RoslynWorkspaceManager?> _getWorkspace;
    private readonly SemanticAnalysisService _semanticService = new();
    private readonly Pen _squigglePen;
    private readonly ToolTip _tooltip;

    private CancellationTokenSource? _diagnosticCts;
    private HoverContext? _hoverContext;
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

        _tooltip = new ToolTip { Placement = PlacementMode.Mouse };

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
        var text = _editor.Text;
        var metadata = _getMetadata();

        var newMarkers = new List<ErrorMarker>();

        // LET structural errors (synchronous — no Roslyn needed)
        var letErrors = LetFormulaValidator.Validate(text);
        foreach (var letError in letErrors)
        {
            newMarkers.Add(new ErrorMarker(letError.StartOffset, letError.Length, letError.Message));
        }

        // Roslyn diagnostics for backtick expressions
        var workspace = _getWorkspace();
        var buildResult = workspace != null
            ? SyntheticDocumentBuilder.BuildForDiagnostics(text, metadata)
            : null;
        if (buildResult != null)
        {
            ct.ThrowIfCancellationRequested();

            var diagnostics = await workspace!.GetDiagnosticsAsync(buildResult.Source, ct);

            ct.ThrowIfCancellationRequested();

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

                var editorOffset = span.Start - exprStart + buildResult.ExpressionStartInEditor;
                var length = Math.Min(span.Length,
                    buildResult.ExpressionLength - (span.Start - exprStart));
                if (length <= 0)
                {
                    length = 1;
                }

                newMarkers.Add(new ErrorMarker(editorOffset, length, diagnostic.GetMessage()));
            }
        }

        _markers = newMarkers;

        // Build semantic model for hover type display
        _hoverContext = BuildHoverContext(text, metadata);

        _editor.TextArea.TextView.InvalidateLayer(Layer);
    }

    private HoverContext? BuildHoverContext(string formulaText, WorkbookMetadata? metadata)
    {
        try
        {
            var expressions = BacktickExtractor.Extract(formulaText);
            if (expressions.Count == 0)
            {
                return null;
            }

            var lastExpr = expressions[^1];
            var expression = lastExpr.Expression;
            var isStatementBlock = expression.TrimStart().StartsWith('{');
            // +1 to skip the opening backtick character
            var expressionStartInEditor = lastExpr.StartIndex + 1;

            // Extract LET bindings
            var letBindings = ExtractLetBindings(formulaText);

            var analysisResult = _semanticService.BuildSemanticModel(
                expression, isStatementBlock, metadata, letBindings);

            return new HoverContext(analysisResult, metadata, expressionStartInEditor, expression.Length);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"BuildHoverContext error: {ex.Message}");
            return null;
        }
    }

    private static IReadOnlyList<(string Name, string Value)>? ExtractLetBindings(string formulaText)
    {
        if (LetFormulaParser.TryParse(formulaText, out var structure) && structure != null)
        {
            return structure.Bindings
                .Select(b => (b.VariableName, b.Value))
                .ToList();
        }

        // Tolerant fallback for incomplete LET formulas
        var letIdx = formulaText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return null;
        }

        var args = LetArgumentSplitter.SplitTolerant(formulaText, letIdx + 4);
        var bindings = new List<(string, string)>();
        for (var i = 0; i + 1 < args.Count; i += 2)
        {
            if (args.Count % 2 == 1 && i == args.Count - 1)
            {
                break;
            }

            bindings.Add((args[i], args[i + 1]));
        }

        return bindings.Count > 0 ? bindings : null;
    }

    private void OnMouseMove(object sender, MouseEventArgs e)
    {
        var pos = _editor.TextArea.TextView.GetPosition(e.GetPosition(_editor.TextArea.TextView));
        if (pos == null)
        {
            _tooltip.IsOpen = false;
            return;
        }

        var offset = _editor.Document.GetOffset(pos.Value.Location);

        // Error tooltips take priority
        var hit = _markers.Find(m => offset >= m.StartOffset && offset < m.StartOffset + m.Length);
        if (hit != null)
        {
            _tooltip.Content = hit.Message;
            _tooltip.IsOpen = true;
            return;
        }

        // Type tooltip on hover over identifiers
        var typeDisplay = GetTypeAtEditorOffset(offset);
        if (typeDisplay != null)
        {
            _tooltip.Content = typeDisplay;
            _tooltip.IsOpen = true;
            return;
        }

        _tooltip.IsOpen = false;
    }

    private string? GetTypeAtEditorOffset(int editorOffset)
    {
        var ctx = _hoverContext;
        if (ctx == null)
        {
            return null;
        }

        // Map editor offset to expression offset
        var expressionOffset = editorOffset - ctx.ExpressionStartInEditor;
        if (expressionOffset < 0 || expressionOffset >= ctx.ExpressionLength)
        {
            return null;
        }

        try
        {
            var type = _semanticService.GetTypeAtOffset(ctx.AnalysisResult, expressionOffset);
            if (type == null || type.TypeKind == TypeKind.Error)
            {
                return null;
            }

            return _semanticService.FormatTypeForDisplay(type, ctx.Metadata);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"GetTypeAtEditorOffset error: {ex.Message}");
            return null;
        }
    }

    private sealed record ErrorMarker(int StartOffset, int Length, string Message);

    private sealed record HoverContext(
        SemanticAnalysisResult AnalysisResult,
        WorkbookMetadata? Metadata,
        int ExpressionStartInEditor,
        int ExpressionLength);
}
