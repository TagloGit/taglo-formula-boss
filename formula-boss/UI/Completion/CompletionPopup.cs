using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;
using ICSharpCode.AvalonEdit.Rendering;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Popup-based completion host that replaces AvalonEdit's <see cref="CompletionWindow" />.
///     Uses <see cref="PlacementMode.Custom" /> with relative coordinates to avoid the DPI
///     mismatch that causes mispositioned completion windows on multi-monitor setups.
///     Hosts AvalonEdit's <see cref="CompletionList" /> directly for filtering, selection,
///     and item rendering.
/// </summary>
internal sealed class CompletionPopup
{
    private static readonly Brush HighlightBrush = new SolidColorBrush(Color.FromRgb(180, 205, 235));
    private readonly Popup _popup;

    private readonly TextArea _textArea;
    private int _endOffset;

    public CompletionPopup(TextArea textArea)
    {
        _textArea = textArea;

        CompletionList = new CompletionList { IsFiltering = true, MaxHeight = 300 };

        var border = new Border
        {
            Background = SystemColors.WindowBrush,
            BorderBrush = SystemColors.ActiveBorderBrush,
            BorderThickness = new Thickness(1),
            MinWidth = 250,
            MaxWidth = 600,
            Child = CompletionList
        };

        _popup = new Popup
        {
            Child = border,
            Placement = PlacementMode.Custom,
            CustomPopupPlacementCallback = PlacePopup,
            PlacementTarget = textArea.TextView,
            StaysOpen = true,
            AllowsTransparency = true
        };

        CompletionList.InsertionRequested += OnInsertionRequested;
    }

    public CompletionList CompletionList { get; }

    /// <summary>
    ///     Offset of the start of the user's typed prefix (i.e. where insertion replaces from).
    ///     Defaults to -1 as a sentinel; if not assigned by the caller before <see cref="Show" />,
    ///     it is initialized to the caret offset (an empty replacement segment).
    /// </summary>
    public int StartOffset { get; set; } = -1;

    public bool CloseWhenCaretAtBeginning { get; set; }

    /// <summary>
    ///     True when the popup is hosting column completions (dot-rewrite or bracket).
    ///     The value-add of these completions is the bracket transformation on commit,
    ///     not filtering — so when AvalonEdit's filter empties the visible list because
    ///     the typed prefix exactly matches an item, the popup stays open with that item
    ///     selected so Enter/Tab still triggers the rewrite.
    /// </summary>
    public bool IsColumnCompletion { get; set; }

    public bool IsOpen => _popup.IsOpen;

    public event EventHandler? Closed;

    public void Show()
    {
        _endOffset = _textArea.Caret.Offset;

        // Defensive default: if caller forgot to set StartOffset, treat the segment
        // as empty (no prefix to replace) rather than [0, caret) which would replace
        // the entire formula on commit.
        if (StartOffset < 0 || StartOffset > _endOffset)
        {
            StartOffset = _endOffset;
        }

        _textArea.Document.Changing += OnDocumentChanging;
        _textArea.Caret.PositionChanged += OnCaretPositionChanged;
        _textArea.PreviewKeyDown += OnPreviewKeyDown;

        var parentWindow = Window.GetWindow(_textArea);
        if (parentWindow != null)
        {
            parentWindow.LocationChanged += OnParentLocationChanged;
        }

        _popup.IsOpen = true;

        // Style the ListBox after the popup opens — CompletionList.ListBox
        // returns null until the control's template is applied in a visual tree.
        var listBox = CompletionList.ListBox;
        if (listBox != null)
        {
            ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
            listBox.Resources[SystemColors.HighlightBrushKey] = HighlightBrush;
            listBox.Resources[SystemColors.HighlightTextBrushKey] = Brushes.Black;
        }

        // Defer focus-loss subscription: creating the Popup's HWND can cause a
        // transient LostKeyboardFocus on the TextArea. A handler dispatched at
        // Normal priority would close the popup before it ever renders.
        _textArea.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
        {
            if (_popup.IsOpen)
            {
                _textArea.LostKeyboardFocus += OnLostFocus;
            }
        }));
    }

    public void Close()
    {
        if (!_popup.IsOpen)
        {
            return;
        }

        _popup.IsOpen = false;

        _textArea.Document.Changing -= OnDocumentChanging;
        _textArea.Caret.PositionChanged -= OnCaretPositionChanged;
        _textArea.PreviewKeyDown -= OnPreviewKeyDown;
        _textArea.LostKeyboardFocus -= OnLostFocus;

        var parentWindow = Window.GetWindow(_textArea);
        if (parentWindow != null)
        {
            parentWindow.LocationChanged -= OnParentLocationChanged;
        }

        Closed?.Invoke(this, EventArgs.Empty);
    }

    private void OnInsertionRequested(object? sender, EventArgs e)
    {
        var item = CompletionList.SelectedItem;
        if (item == null)
        {
            Close();
            return;
        }

        // Close before completing — AvalonEdit's design expects this order
        var startOffset = StartOffset;
        var endOffset = _endOffset;
        Close();

        // Defensive: a corrupted [start, end) range (e.g. start=0 with non-empty doc when
        // the popup opened with no prefix) would otherwise replace text outside the typed
        // prefix span. Refuse to commit a non-empty completion with an invalid range.
        if (startOffset < 0 || endOffset < startOffset || endOffset > _textArea.Document.TextLength)
        {
            return;
        }

        var segment = new AnchorSegment(_textArea.Document, startOffset, endOffset - startOffset);
        item.Complete(_textArea, segment, e);
    }

    private void OnDocumentChanging(object? sender, DocumentChangeEventArgs e)
    {
        if (e.Offset >= StartOffset && e.Offset <= _endOffset)
        {
            _endOffset += e.InsertionLength - e.RemovalLength;
        }
        else if (e.Offset < StartOffset)
        {
            // If removal reaches into our region, close
            if (e.Offset + e.RemovalLength > StartOffset)
            {
                Close();
                return;
            }

            var shift = e.InsertionLength - e.RemovalLength;
            StartOffset += shift;
            _endOffset += shift;
        }
    }

    private void OnCaretPositionChanged(object? sender, EventArgs e)
    {
        var caretOffset = _textArea.Caret.Offset;

        if (caretOffset < StartOffset || (caretOffset == StartOffset && CloseWhenCaretAtBeginning))
        {
            Close();
            return;
        }

        if (caretOffset > _endOffset)
        {
            Close();
            return;
        }

        // Update filtering
        var text = _textArea.Document.GetText(StartOffset, caretOffset - StartOffset);
        CompletionList.SelectItem(text);

        // If filtering left no visible items, close — except for column completion,
        // where the typed prefix exactly matching an item (e.g. column "1" + user types "1")
        // can cause AvalonEdit to drop the item from the visible list. In that case keep
        // the popup open with the matching item selected so Enter/Tab still rewrites to
        // bracket syntax.
        if (CompletionList.ListBox?.Items.Count == 0)
        {
            if (IsColumnCompletion && TryReselectExactColumnMatch(text))
            {
                _popup.HorizontalOffset = 0;
                return;
            }

            Close();
            return;
        }

        // Force popup to reposition to follow the caret
        _popup.HorizontalOffset = 0;
    }

    /// <summary>
    ///     When AvalonEdit's filter has emptied the visible list, look for a column
    ///     completion item whose text exactly matches the typed prefix. If found, repopulate
    ///     the listbox with just that item and select it so a subsequent Enter/Tab triggers
    ///     <see cref="OnInsertionRequested" />.
    /// </summary>
    private bool TryReselectExactColumnMatch(string typedPrefix)
    {
        if (string.IsNullOrEmpty(typedPrefix) || CompletionList.ListBox == null)
        {
            return false;
        }

        ICompletionData? match = null;
        foreach (var data in CompletionList.CompletionData)
        {
            if (data.Text.Equals(typedPrefix, StringComparison.OrdinalIgnoreCase))
            {
                match = data;
                break;
            }
        }

        if (match == null)
        {
            return false;
        }

        CompletionList.ListBox.ItemsSource = new ObservableCollection<ICompletionData> { match };
        CompletionList.SelectedItem = match;
        return true;
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Up:
            case Key.Down:
            case Key.PageUp:
            case Key.PageDown:
            case Key.Home:
            case Key.End:
                // Route navigation keys to the completion list
                CompletionList.HandleKey(e);
                break;

            case Key.Tab:
            case Key.Enter:
                CompletionList.RequestInsertion(e);
                e.Handled = true;
                break;

            case Key.Escape:
                Close();
                e.Handled = true;
                break;
        }
    }

    private void OnLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        // Async dispatch to check if focus has truly left the text area
        _textArea.Dispatcher.BeginInvoke(new Action(() =>
        {
            if (!_textArea.IsKeyboardFocusWithin)
            {
                Close();
            }
        }));
    }

    private void OnParentLocationChanged(object? sender, EventArgs e) =>
        // Force WPF to recalculate popup placement when the parent window moves
        _popup.HorizontalOffset = 0;

    /// <summary>
    ///     DPI-safe placement callback using relative coordinates.
    ///     Subtracts PointToScreen(origin) from PointToScreen(caret), producing
    ///     DPI-neutral offsets — the same pattern used by <see cref="SignatureHelpPopup" />.
    /// </summary>
    private CustomPopupPlacement[] PlacePopup(Size popupSize, Size targetSize, Point offset)
    {
        var textView = _textArea.TextView;
        var caretPos = textView.GetVisualPosition(
            _textArea.Caret.Position, VisualYPosition.LineBottom);
        var caretScreen = textView.PointToScreen(caretPos);
        var targetScreen = textView.PointToScreen(new Point(0, 0));

        var x = caretScreen.X - targetScreen.X;
        var y = caretScreen.Y - targetScreen.Y + 2;

        // If not enough room below, place above the caret line
        var caretTopPos = textView.GetVisualPosition(
            _textArea.Caret.Position, VisualYPosition.LineTop);
        var caretTopScreen = textView.PointToScreen(caretTopPos);

        // Estimate available space below using target size as proxy for visible area
        var spaceBelow = targetSize.Height - (caretScreen.Y - targetScreen.Y);
        if (spaceBelow < popupSize.Height)
        {
            y = caretTopScreen.Y - targetScreen.Y - popupSize.Height - 2;
        }

        return [new CustomPopupPlacement(new Point(x, y), PopupPrimaryAxis.Horizontal)];
    }
}
