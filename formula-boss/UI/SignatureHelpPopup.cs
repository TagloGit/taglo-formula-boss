using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;

using FormulaBoss.UI.Completion;

using ICSharpCode.AvalonEdit;

namespace FormulaBoss.UI;

/// <summary>
///     Displays VS-style signature help (parameter info) near the caret.
///     Shows method signature with active parameter bold, overload counter,
///     method summary, and active parameter description.
/// </summary>
internal sealed class SignatureHelpPopup
{
    private static readonly Brush BackgroundBrush = new SolidColorBrush(Color.FromRgb(0xF6, 0xF6, 0xF6));
    private static readonly Brush BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCE, 0xDB));

    private readonly TextEditor _editor;
    private readonly Popup _popup;
    private readonly TextBlock _overloadCounter;
    private readonly TextBlock _signatureLine;
    private readonly TextBlock _summaryText;
    private readonly TextBlock _parameterText;

    private SignatureHelpModel? _model;
    private int _selectedOverloadIndex;

    public SignatureHelpPopup(TextEditor editor)
    {
        _editor = editor;

        _overloadCounter = new TextBlock
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 12,
            Foreground = Brushes.Gray,
            Margin = new Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

        _signatureLine = new TextBlock
        {
            FontFamily = new FontFamily("Consolas"),
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center
        };

        var headerPanel = new DockPanel { Margin = new Thickness(0, 0, 0, 2) };
        DockPanel.SetDock(_overloadCounter, Dock.Left);
        headerPanel.Children.Add(_overloadCounter);
        headerPanel.Children.Add(_signatureLine);

        _summaryText = new TextBlock
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 12,
            Foreground = Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 2, 0, 0)
        };

        _parameterText = new TextBlock
        {
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 2, 0, 0)
        };

        var panel = new StackPanel { Margin = new Thickness(6, 4, 6, 4) };
        panel.Children.Add(headerPanel);
        panel.Children.Add(_summaryText);
        panel.Children.Add(_parameterText);

        var border = new Border
        {
            Background = BackgroundBrush,
            BorderBrush = BorderBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(2),
            MaxWidth = 500,
            Child = panel
        };

        _popup = new Popup
        {
            Child = border,
            Placement = PlacementMode.Custom,
            CustomPopupPlacementCallback = PlacePopup,
            PlacementTarget = editor.TextArea.TextView,
            StaysOpen = true,
            AllowsTransparency = true
        };
    }

    public bool IsVisible => _popup.IsOpen;

    public void Update(SignatureHelpModel model)
    {
        _model = model;

        // Preserve user's overload selection if still valid, otherwise use Roslyn's suggestion
        if (_selectedOverloadIndex >= model.Overloads.Count)
        {
            _selectedOverloadIndex = model.ActiveOverloadIndex;
        }

        Render();
        _popup.IsOpen = true;
    }

    public void Hide()
    {
        _popup.IsOpen = false;
        _model = null;
        _selectedOverloadIndex = 0;
    }

    /// <summary>
    ///     Cycles to the next overload. Returns true if consumed (multiple overloads exist).
    /// </summary>
    public bool NextOverload()
    {
        if (_model == null || _model.Overloads.Count <= 1)
        {
            return false;
        }

        _selectedOverloadIndex = (_selectedOverloadIndex + 1) % _model.Overloads.Count;
        Render();
        return true;
    }

    /// <summary>
    ///     Cycles to the previous overload. Returns true if consumed (multiple overloads exist).
    /// </summary>
    public bool PreviousOverload()
    {
        if (_model == null || _model.Overloads.Count <= 1)
        {
            return false;
        }

        _selectedOverloadIndex = (_selectedOverloadIndex - 1 + _model.Overloads.Count) % _model.Overloads.Count;
        Render();
        return true;
    }

    private void Render()
    {
        if (_model == null)
        {
            return;
        }

        var overload = _model.Overloads[_selectedOverloadIndex];
        var activeParam = _model.ActiveParameterIndex;

        // Overload counter
        if (_model.Overloads.Count > 1)
        {
            _overloadCounter.Text = $"\u25B2 {_selectedOverloadIndex + 1} of {_model.Overloads.Count} \u25BC";
            _overloadCounter.Visibility = Visibility.Visible;
        }
        else
        {
            _overloadCounter.Visibility = Visibility.Collapsed;
        }

        // Signature line with active parameter bold
        _signatureLine.Inlines.Clear();
        _signatureLine.Inlines.Add(new Run(overload.MethodName + "("));

        for (var i = 0; i < overload.Parameters.Count; i++)
        {
            if (i > 0)
            {
                _signatureLine.Inlines.Add(new Run(", "));
            }

            var param = overload.Parameters[i];
            var paramText = $"{param.Type} {param.Name}";
            var run = new Run(paramText);

            if (i == activeParam)
            {
                run.FontWeight = FontWeights.Bold;
            }

            _signatureLine.Inlines.Add(run);
        }

        _signatureLine.Inlines.Add(new Run("): " + overload.ReturnType));

        // Summary
        if (!string.IsNullOrEmpty(overload.Summary))
        {
            _summaryText.Text = overload.Summary;
            _summaryText.Visibility = Visibility.Visible;
        }
        else
        {
            _summaryText.Visibility = Visibility.Collapsed;
        }

        // Active parameter description
        if (activeParam >= 0 && activeParam < overload.Parameters.Count)
        {
            var param = overload.Parameters[activeParam];
            if (!string.IsNullOrEmpty(param.Description))
            {
                _parameterText.Inlines.Clear();
                _parameterText.Inlines.Add(new Run(param.Name + ": ") { FontWeight = FontWeights.SemiBold });
                _parameterText.Inlines.Add(new Run(param.Description));
                _parameterText.Visibility = Visibility.Visible;
            }
            else
            {
                _parameterText.Visibility = Visibility.Collapsed;
            }
        }
        else
        {
            _parameterText.Visibility = Visibility.Collapsed;
        }

        // Reposition to current caret
        _popup.HorizontalOffset = 0; // Force WPF to recalculate placement
    }

    private CustomPopupPlacement[] PlacePopup(Size popupSize, Size targetSize, Point offset)
    {
        var textView = _editor.TextArea.TextView;
        var caretPos = textView.GetVisualPosition(
            _editor.TextArea.Caret.Position,
            ICSharpCode.AvalonEdit.Rendering.VisualYPosition.LineTop);
        var caretScreen = textView.PointToScreen(caretPos);
        var targetScreen = textView.PointToScreen(new Point(0, 0));

        var x = caretScreen.X - targetScreen.X;
        var y = caretScreen.Y - targetScreen.Y - popupSize.Height - 2;

        // If not enough room above, place below the line
        if (y < 0)
        {
            var lineBottom = textView.GetVisualPosition(
                _editor.TextArea.Caret.Position,
                ICSharpCode.AvalonEdit.Rendering.VisualYPosition.LineBottom);
            y = textView.PointToScreen(lineBottom).Y - targetScreen.Y + 2;
        }

        return [new CustomPopupPlacement(new Point(x, y), PopupPrimaryAxis.Horizontal)];
    }
}
