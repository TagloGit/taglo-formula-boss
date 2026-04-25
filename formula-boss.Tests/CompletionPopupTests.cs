using FormulaBoss.UI;
using FormulaBoss.UI.Completion;

using ICSharpCode.AvalonEdit;

using Xunit;

namespace FormulaBoss.Tests;

public class CompletionPopupTests
{
    private static void RunOnSta(Action action)
    {
        Exception? caught = null;
        var thread = new Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex) { caught = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (caught != null)
        {
            throw caught;
        }
    }

    private static (TextEditor editor, CompletionPopup popup) CreatePopup(string text, int caretOffset)
    {
        var editor = new TextEditor { Text = text, CaretOffset = caretOffset };
        var popup = new CompletionPopup(editor.TextArea);
        return (editor, popup);
    }

    [Fact]
    public void Show_with_unset_StartOffset_initializes_to_caret_offset() => RunOnSta(() =>
    {
        // Repro for issue #324 1a: caller forgets to set StartOffset (e.g. wordLength == 0
        // path before the fix). Show() must default StartOffset to the caret offset so the
        // would-be replacement segment is empty rather than [0, caret).
        var (_, popup) = CreatePopup("var myRow = t.Rows.First(); return myRow.", 41);
        try
        {
            popup.Show();
            Assert.Equal(41, popup.StartOffset);
        }
        finally
        {
            popup.Close();
        }
    });

    [Fact]
    public void Show_with_explicit_StartOffset_preserves_value() => RunOnSta(() =>
    {
        // Caller has computed wordLength > 0 and set StartOffset back to the start of the
        // typed prefix — Show must not overwrite it.
        var (_, popup) = CreatePopup("r.Pl", 4);
        popup.StartOffset = 2;
        try
        {
            popup.Show();
            Assert.Equal(2, popup.StartOffset);
        }
        finally
        {
            popup.Close();
        }
    });

    [Fact]
    public void Show_with_StartOffset_beyond_caret_is_clamped_to_caret() => RunOnSta(() =>
    {
        // Defensive: a StartOffset beyond the caret would yield a negative segment length.
        // Show should clamp it back to the caret offset.
        var (_, popup) = CreatePopup("hello", 3);
        popup.StartOffset = 100;
        try
        {
            popup.Show();
            Assert.Equal(3, popup.StartOffset);
        }
        finally
        {
            popup.Close();
        }
    });

    [Fact]
    public void Show_with_zero_StartOffset_at_zero_caret_is_valid() => RunOnSta(() =>
    {
        // StartOffset == 0 is legitimate when the caret is also at 0 (empty document).
        var (_, popup) = CreatePopup("", 0);
        popup.StartOffset = 0;
        try
        {
            popup.Show();
            Assert.Equal(0, popup.StartOffset);
        }
        finally
        {
            popup.Close();
        }
    });

    [Fact]
    public void Insertion_with_empty_segment_at_dot_offset_replaces_only_dot_via_ColumnCompletion() => RunOnSta(() =>
    {
        // End-to-end coverage of issue #324 1a: with the bug, StartOffset was 0 and Enter
        // would replace the entire formula. With the fix, StartOffset == caret, so the
        // ColumnCompletionData rewrites only the dot to bracket syntax.
        var formula = "var myRow = t.Rows.First(); return myRow.";
        var (editor, popup) = CreatePopup(formula, formula.Length);
        var item = new ColumnCompletionData("Player", "Column name", isBracketContext: false);
        popup.CompletionList.CompletionData.Add(item);
        popup.CompletionList.SelectedItem = item;

        try
        {
            popup.Show();
            popup.CompletionList.RequestInsertion(EventArgs.Empty);
        }
        finally
        {
            // RequestInsertion calls Close internally; avoid double-close
            if (popup.IsOpen)
            {
                popup.Close();
            }
        }

        Assert.Equal("var myRow = t.Rows.First(); return myRow[\"Player\"]", editor.Text);
    });

    [Fact]
    public void Insertion_after_typed_prefix_replaces_dot_and_prefix() => RunOnSta(() =>
    {
        // Issue #324 acceptance: typing `myRow.Player` then Enter rewrites to myRow["Player"]
        // and does not touch the surrounding text.
        var formula = "var myRow = t.Rows.First(); return myRow.";
        var (editor, popup) = CreatePopup(formula, formula.Length);
        var item = new ColumnCompletionData("Player", "Column name", isBracketContext: false);
        popup.CompletionList.CompletionData.Add(item);
        popup.CompletionList.SelectedItem = item;

        try
        {
            popup.Show();

            // User types "Player" after the dot — document inserts shift _endOffset
            editor.Document.Insert(editor.CaretOffset, "Player");

            popup.CompletionList.RequestInsertion(EventArgs.Empty);
        }
        finally
        {
            if (popup.IsOpen)
            {
                popup.Close();
            }
        }

        Assert.Equal("var myRow = t.Rows.First(); return myRow[\"Player\"]", editor.Text);
    });

    [Fact]
    public void Insertion_in_column_mode_with_single_char_column_uses_bracket_rewrite() => RunOnSta(() =>
    {
        // Repro for issue #324 1b: column "1" / typed "1". Even if AvalonEdit's filter
        // empties the visible list, the popup must stay open in column mode and Enter
        // must still trigger the bracket rewrite.
        var formula = "r.";
        var (editor, popup) = CreatePopup(formula, formula.Length);
        var col1 = new ColumnCompletionData("1", "Column name", isBracketContext: false);
        var col2 = new ColumnCompletionData("2", "Column name", isBracketContext: false);
        var col3 = new ColumnCompletionData("3", "Column name", isBracketContext: false);
        popup.CompletionList.CompletionData.Add(col1);
        popup.CompletionList.CompletionData.Add(col2);
        popup.CompletionList.CompletionData.Add(col3);
        popup.CompletionList.SelectedItem = col1;
        popup.IsColumnCompletion = true;

        try
        {
            popup.Show();

            // User types "1"
            editor.Document.Insert(editor.CaretOffset, "1");

            popup.CompletionList.RequestInsertion(EventArgs.Empty);
        }
        finally
        {
            if (popup.IsOpen)
            {
                popup.Close();
            }
        }

        Assert.Equal("r[\"1\"]", editor.Text);
    });
}
