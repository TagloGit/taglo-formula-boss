using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

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

    private readonly TextArea _textArea;
    private readonly Popup _popup;
    private readonly Border _border;
    private int _endOffset;

    public CompletionPopup(TextArea textArea)
    {
        _textArea = textArea;

        CompletionList = new CompletionList { IsFiltering = true };

        // Style the list box
        var listBox = CompletionList.ListBox;
        ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
        listBox.Resources[SystemColors.HighlightBrushKey] = HighlightBrush;
        listBox.Resources[SystemColors.HighlightTextBrushKey] = Brushes.Black;
        listBox.MaxHeight = 300;

        _border = new Border
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
            Child = _border,
            Placement = PlacementMode.Custom,
            CustomPopupPlacementCallback = PlacePopup,
            PlacementTarget = textArea.TextView,
            StaysOpen = true,
            AllowsTransparency = true
        };

        CompletionList.InsertionRequested += OnInsertionRequested;
    }

    public CompletionList CompletionList { get; }

    public int StartOffset { get; set; }

    public bool CloseWhenCaretAtBeginning { get; set; }

    public bool IsOpen => _popup.IsOpen;

    public event EventHandler? Closed;

    public void Show()
    {
        _endOffset = _textArea.Caret.Offset;

        _textArea.Document.Changing += OnDocumentChanging;
        _textArea.Caret.PositionChanged += OnCaretPositionChanged;
        _textArea.PreviewKeyDown += OnPreviewKeyDown;
        _textArea.LostKeyboardFocus += OnLostFocus;

        var parentWindow = Window.GetWindow(_textArea);
        if (parentWindow != null)
        {
            parentWindow.LocationChanged += OnParentLocationChanged;
        }

        _popup.IsOpen = true;
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
            var shift = e.InsertionLength - e.RemovalLength;
            // If removal reaches into our region, close
            if (e.Offset + e.RemovalLength > StartOffset)
            {
                Close();
                return;
            }

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

        // If filtering left no visible items, close
        if (CompletionList.ListBox.Items.Count == 0)
        {
            Close();
            return;
        }

        // Force popup to reposition to follow the caret
        _popup.HorizontalOffset = 0;
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

    private void OnParentLocationChanged(object? sender, EventArgs e)
    {
        // Force WPF to recalculate popup placement when the parent window moves
        _popup.HorizontalOffset = 0;
    }

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
