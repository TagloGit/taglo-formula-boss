using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;

namespace FormulaBoss.UI;

/// <summary>
///     A single item in the completion popup.
/// </summary>
public class CompletionData : ICompletionData
{
    public CompletionData(string text, string? description = null)
    {
        Text = text;
        DescriptionText = description;
    }

    public string? DescriptionText { get; }

    public string Text { get; }

    public object Content
    {
        get
        {
            if (DescriptionText == null)
            {
                return Text;
            }

            return new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Children =
                {
                    new TextBlock { Text = Text, FontWeight = FontWeights.Medium },
                    new TextBlock
                    {
                        Text = DescriptionText,
                        Margin = new Thickness(32, 0, 0, 0),
                        Foreground = Brushes.Gray,
                        FontStyle = FontStyles.Italic
                    }
                }
            };
        }
    }

    // Return null to suppress AvalonEdit's built-in tooltip.
    public object? Description => null;

    public double Priority { get; init; }
    public ImageSource? Image => null;

    public virtual void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs) =>
        textArea.Document.Replace(completionSegment, Text);
}

/// <summary>
///     Completion item for row column names with context-aware insertion.
///     After a dot: rewrites to bracket syntax if the name contains spaces.
///     After a bracket: appends closing bracket.
/// </summary>
internal sealed class ColumnCompletionData : CompletionData
{
    private readonly bool _isBracketContext;

    public ColumnCompletionData(string text, string? description, bool isBracketContext)
        : base(text, description)
    {
        _isBracketContext = isBracketContext;
    }

    public override void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
    {
        var doc = textArea.Document;
        var hasSpace = Text.Contains(' ');

        if (_isBracketContext)
        {
            // Inside r[ — insert column name + ]
            doc.Replace(completionSegment, Text + "]");
        }
        else if (hasSpace)
        {
            // After r. — rewrite the dot to bracket syntax: r[Column Name]
            // The dot is the character immediately before the completion segment
            var dotOffset = completionSegment.Offset - 1;
            if (dotOffset >= 0 && doc.GetCharAt(dotOffset) == '.')
            {
                doc.Replace(dotOffset, completionSegment.EndOffset - dotOffset, "[" + Text + "]");
            }
            else
            {
                // Fallback: just insert as bracket syntax
                doc.Replace(completionSegment, "[" + Text + "]");
            }
        }
        else
        {
            // No space, dot context — normal insertion
            doc.Replace(completionSegment, Text);
        }
    }
}
