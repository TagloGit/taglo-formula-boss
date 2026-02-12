using System.Windows.Media;

using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;

namespace FormulaBoss.UI;

/// <summary>
/// A single item in the completion popup.
/// </summary>
public class CompletionData : ICompletionData
{
    public CompletionData(string text, string? description = null)
    {
        Text = text;
        Description = description;
    }

    public string Text { get; }
    public object Content => Text;
    public object? Description { get; }
    public double Priority { get; init; }
    public ImageSource? Image => null;

    public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
    {
        textArea.Document.Replace(completionSegment, Text);
    }
}
