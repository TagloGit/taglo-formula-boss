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

    public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs) =>
        textArea.Document.Replace(completionSegment, Text);
}
