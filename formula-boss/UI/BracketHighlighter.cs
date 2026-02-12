using System.Windows.Media;

using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;

namespace FormulaBoss.UI;

/// <summary>
///     Highlights matching bracket pairs in the editor.
///     Searches for matching brackets when the caret moves, and renders
///     a background highlight on both the opening and closing bracket.
/// </summary>
internal sealed class BracketHighlighter : IBackgroundRenderer
{
    private static readonly Dictionary<char, char> OpenToClose = new() { { '(', ')' }, { '[', ']' }, { '{', '}' } };

    private static readonly Dictionary<char, char> CloseToOpen = new() { { ')', '(' }, { ']', '[' }, { '}', '{' } };

    private readonly Brush _backgroundBrush;
    private readonly Pen _borderPen;

    private readonly TextEditor _editor;
    private int _closeOffset = -1;
    private int _openOffset = -1;

    public BracketHighlighter(TextEditor editor)
    {
        _editor = editor;

        var borderColor = Color.FromArgb(180, 0, 100, 200);
        _borderPen = new Pen(new SolidColorBrush(borderColor), 1);
        _borderPen.Freeze();

        var bgColor = Color.FromArgb(40, 0, 100, 200);
        _backgroundBrush = new SolidColorBrush(bgColor);
        _backgroundBrush.Freeze();

        editor.TextArea.TextView.BackgroundRenderers.Add(this);
        editor.TextArea.Caret.PositionChanged += (_, _) => UpdateHighlight();
    }

    public KnownLayer Layer => KnownLayer.Selection;

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (_openOffset < 0 || _closeOffset < 0)
        {
            return;
        }

        var builder = new BackgroundGeometryBuilder { CornerRadius = 1 };

        builder.AddSegment(textView, new TextSegment { StartOffset = _openOffset, Length = 1 });
        builder.CloseFigure();
        builder.AddSegment(textView, new TextSegment { StartOffset = _closeOffset, Length = 1 });

        var geometry = builder.CreateGeometry();
        if (geometry != null)
        {
            drawingContext.DrawGeometry(_backgroundBrush, _borderPen, geometry);
        }
    }

    private void UpdateHighlight()
    {
        var oldOpen = _openOffset;
        var oldClose = _closeOffset;

        _openOffset = -1;
        _closeOffset = -1;

        var doc = _editor.Document;
        var offset = _editor.CaretOffset;

        // Check character before caret
        if (offset > 0)
        {
            TryMatch(doc, offset - 1);
        }

        // Check character at caret
        if (_openOffset < 0 && offset < doc.TextLength)
        {
            TryMatch(doc, offset);
        }

        if (oldOpen != _openOffset || oldClose != _closeOffset)
        {
            _editor.TextArea.TextView.InvalidateLayer(Layer);
        }
    }

    private void TryMatch(TextDocument doc, int offset)
    {
        var ch = doc.GetCharAt(offset);

        if (OpenToClose.TryGetValue(ch, out var expectedClose))
        {
            var match = FindMatchingForward(doc, offset + 1, ch, expectedClose);
            if (match >= 0)
            {
                _openOffset = offset;
                _closeOffset = match;
            }
        }
        else if (CloseToOpen.TryGetValue(ch, out var expectedOpen))
        {
            var match = FindMatchingBackward(doc, offset - 1, expectedOpen, ch);
            if (match >= 0)
            {
                _openOffset = match;
                _closeOffset = offset;
            }
        }
    }

    private static int FindMatchingForward(TextDocument doc, int start, char open, char close)
    {
        var depth = 1;
        var inString = false;
        var inChar = false;

        for (var i = start; i < doc.TextLength; i++)
        {
            var ch = doc.GetCharAt(i);

            // Skip escaped characters
            if (i > 0 && doc.GetCharAt(i - 1) == '\\')
            {
                continue;
            }

            if (ch == '"' && !inChar)
            {
                inString = !inString;
            }
            else if (ch == '\'' && !inString)
            {
                inChar = !inChar;
            }

            if (inString || inChar)
            {
                continue;
            }

            if (ch == open)
            {
                depth++;
            }
            else if (ch == close)
            {
                depth--;
            }

            if (depth == 0)
            {
                return i;
            }
        }

        return -1;
    }

    private static int FindMatchingBackward(TextDocument doc, int start, char open, char close)
    {
        var depth = 1;
        var inString = false;
        var inChar = false;

        for (var i = start; i >= 0; i--)
        {
            var ch = doc.GetCharAt(i);

            // Skip escaped characters
            if (i > 0 && doc.GetCharAt(i - 1) == '\\')
            {
                continue;
            }

            if (ch == '"' && !inChar)
            {
                inString = !inString;
            }
            else if (ch == '\'' && !inString)
            {
                inChar = !inChar;
            }

            if (inString || inChar)
            {
                continue;
            }

            if (ch == close)
            {
                depth++;
            }
            else if (ch == open)
            {
                depth--;
            }

            if (depth == 0)
            {
                return i;
            }
        }

        return -1;
    }
}
