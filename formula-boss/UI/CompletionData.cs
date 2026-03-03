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
///     After a dot: always rewrites to bracket syntax ["Column Name"].
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
        var fullText = doc.Text;
        var (newText, replaceOffset, replaceLength) = ComputeInsertion(
            Text, _isBracketContext, fullText, completionSegment.Offset, completionSegment.Length);
        doc.Replace(replaceOffset, replaceLength, newText);
    }

    /// <summary>
    ///     Computes the text to insert and the replacement range for a column completion.
    ///     Separated from the AvalonEdit <see cref="TextArea"/> for testability.
    /// </summary>
    /// <returns>(text to insert, start offset of replacement, length to replace)</returns>
    internal static (string NewText, int ReplaceOffset, int ReplaceLength) ComputeInsertion(
        string columnName, bool isBracketContext, string documentText, int segmentOffset, int segmentLength)
    {
        var quoted = $"\"{columnName}\"";
        var segmentEnd = segmentOffset + segmentLength;

        if (isBracketContext)
        {
            // Check if the auto-closer already placed a ] after the segment
            var hasClosingBracket = segmentEnd < documentText.Length && documentText[segmentEnd] == ']';

            if (hasClosingBracket)
            {
                // Replace segment + the existing ']' so we don't double it
                return (quoted + "]", segmentOffset, segmentLength + 1);
            }

            return (quoted + "]", segmentOffset, segmentLength);
        }

        // Dot context: rewrite the dot to bracket syntax
        var dotOffset = segmentOffset - 1;
        if (dotOffset >= 0 && documentText[dotOffset] == '.')
        {
            return ("[" + quoted + "]", dotOffset, segmentEnd - dotOffset);
        }

        return ("[" + quoted + "]", segmentOffset, segmentLength);
    }
}
