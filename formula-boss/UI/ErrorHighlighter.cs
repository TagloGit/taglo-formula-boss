using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;

using FormulaBoss.Interception;
using FormulaBoss.Parsing;

namespace FormulaBoss.UI;

/// <summary>
///     Highlights parse errors in backtick expressions with red squiggly underlines.
///     Shows tooltip messages on mouse hover over error regions.
/// </summary>
internal sealed class ErrorHighlighter : IBackgroundRenderer
{
    private readonly TextEditor _editor;
    private readonly DispatcherTimer _debounceTimer;
    private readonly Pen _squigglePen;
    private readonly ToolTip _tooltip;

    private List<ErrorMarker> _markers = [];

    public ErrorHighlighter(TextEditor editor)
    {
        _editor = editor;

        var red = Color.FromRgb(0xFF, 0x30, 0x30);
        _squigglePen = new Pen(new SolidColorBrush(red), 1.2) { DashStyle = DashStyles.Dot };
        _squigglePen.Freeze();

        _tooltip = new ToolTip { Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse };

        _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
        _debounceTimer.Tick += (_, _) =>
        {
            _debounceTimer.Stop();
            UpdateErrors();
        };

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

    private void UpdateErrors()
    {
        var text = _editor.Text;
        var newMarkers = new List<ErrorMarker>();

        var expressions = BacktickExtractor.Extract(text);
        foreach (var expr in expressions)
        {
            // Lex and parse the backtick expression
            var lexer = new Lexer(expr.Expression);
            var tokens = lexer.ScanTokens();

            // Collect lexer errors (tokens with TokenType.Error)
            foreach (var token in tokens)
            {
                if (token.Type == TokenType.Error)
                {
                    // +1 skips the opening backtick
                    var offset = expr.StartIndex + 1 + token.Position;
                    newMarkers.Add(new ErrorMarker(offset, Math.Max(1, token.Lexeme.Length), token.Lexeme));
                }
            }

            // Collect parser errors
            var parser = new Parser(tokens, expr.Expression);
            parser.Parse();
            foreach (var (position, length, message) in parser.StructuredErrors)
            {
                var offset = expr.StartIndex + 1 + position;
                newMarkers.Add(new ErrorMarker(offset, length, message));
            }
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
