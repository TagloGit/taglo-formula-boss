using System.Diagnostics;
using System.Windows.Interop;
using System.Windows.Threading;

using ExcelDna.Integration;

using FormulaBoss.Interception;
using FormulaBoss.UI;
using FormulaBoss.UI.Animation;

namespace FormulaBoss.Commands;

/// <summary>
///     Command handler for showing the floating formula editor.
///     Triggered by Ctrl+Shift+` keyboard shortcut.
/// </summary>
public static class ShowFloatingEditorCommand
{
    private static FloatingEditorWindow? _window;
    private static Dispatcher? _windowDispatcher;
    private static Thread? _windowThread;
    private static dynamic? _app;

    // Target cell captured at open time so Apply writes to the correct cell
    // even if the user clicks elsewhere while the editor is open
    private static dynamic? _targetWorksheet;
    private static string? _targetAddress;

    // Cell screen position captured at open time for animation placement
    private static int _targetCellScreenLeft;
    private static int _targetCellScreenTop;

    /// <summary>
    ///     Initializes the command with the Excel application reference.
    ///     Must be called during add-in initialization.
    /// </summary>
    public static void Initialize(dynamic app) => _app = app;

    /// <summary>
    ///     Cleans up the floating editor window.
    ///     Should be called during add-in shutdown.
    /// </summary>
    public static void Cleanup()
    {
        _windowDispatcher?.InvokeShutdown();
        _windowDispatcher = null;

        _windowThread = null;
        _window = null;
        _app = null;
        _targetWorksheet = null;
        _targetAddress = null;
    }

    /// <summary>
    ///     Opens the floating formula editor with context-aware content based on cell state.
    ///     If editor is already visible, hides it (toggle).
    /// </summary>
    [ExcelCommand]
    public static void ShowFloatingEditor()
    {
        try
        {
            _app ??= ExcelDnaUtil.Application;

            var cell = _app.ActiveCell;
            var currentAddress = cell.Address as string;
            var editorContent = DetectEditorContent(cell);
            var excelHwnd = new IntPtr(_app.Hwnd);

            // Capture cell screen position for animation placement
            CaptureTargetCellPosition(cell);

            EnsureWindowThread();

            _windowDispatcher?.Invoke(() =>
            {
                if (_window == null)
                {
                    return;
                }

                if (_window.IsVisible)
                {
                    if (currentAddress == _targetAddress)
                    {
                        // Same cell — toggle off
                        _window.Hide();
                    }
                    else
                    {
                        // Different cell — update content and target
                        _targetWorksheet = cell.Worksheet;
                        _targetAddress = currentAddress;
                        _window.FormulaText = editorContent;

                        var wpfHwnd = new WindowInteropHelper(_window).EnsureHandle();
                        WindowPositioner.CenterOnExcel(excelHwnd, wpfHwnd);
                    }
                }
                else
                {
                    // Capture target cell at open time
                    _targetWorksheet = cell.Worksheet;
                    _targetAddress = currentAddress;
                    _window.FormulaText = editorContent;

                    var wpfHwnd = new WindowInteropHelper(_window).EnsureHandle();
                    WindowPositioner.CenterOnExcel(excelHwnd, wpfHwnd);

                    _window.Show();
                    _window.Activate();
                }
            });
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ShowFloatingEditor error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Detects the appropriate editor content based on the cell's current state.
    /// </summary>
    private static string DetectEditorContent(dynamic cell)
    {
        var formula = cell.Formula2 as string ?? cell.Formula as string ?? "";

        // Empty cell — pre-fill with = so the user can start typing immediately
        if (string.IsNullOrEmpty(formula))
        {
            return "=";
        }

        // Unprocessed FB formula: quote-prefixed text (user typed '=... with backticks)
        var prefix = cell.PrefixCharacter as string;
        if (prefix == "'")
        {
            return cell.Value2 as string ?? "";
        }

        // Processed FB LET formula: contains _src_ variables from pipeline
        if (LetFormulaReconstructor.IsProcessedFormulaBossLet(formula))
        {
            if (LetFormulaReconstructor.TryReconstruct(formula, out var editableFormula) &&
                editableFormula != null)
            {
                // TryReconstruct adds a leading ' for text storage — strip it for the editor
                return editableFormula.StartsWith('\'') ? editableFormula[1..] : editableFormula;
            }
        }

        // Processed basic FB / non-FB formula: show formula as-is
        return formula;
    }

    /// <summary>
    ///     Captures the target cell's screen position using Excel's PointsToScreenPixels.
    /// </summary>
    private static void CaptureTargetCellPosition(dynamic cell)
    {
        try
        {
            var window = _app!.ActiveWindow;
            _targetCellScreenLeft = (int)(double)window.PointsToScreenPixelsX(cell.Left);
            _targetCellScreenTop = (int)(double)window.PointsToScreenPixelsY(cell.Top);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"CaptureTargetCellPosition error: {ex.Message}");
            _targetCellScreenLeft = 0;
            _targetCellScreenTop = 0;
        }
    }

    private static void EnsureWindowThread()
    {
        if (_windowThread != null && _windowThread.IsAlive)
        {
            return;
        }

        var readyEvent = new ManualResetEventSlim(false);

        _windowThread = new Thread(() =>
        {
            NativeMethods.SetThreadDpiAwarenessContext(
                NativeMethods.DpiAwarenessContextPerMonitorAwareV2);

            _window = new FloatingEditorWindow();
            _window.FormulaApplied += OnFormulaApplied;
            _windowDispatcher = Dispatcher.CurrentDispatcher;
            readyEvent.Set();

            Dispatcher.Run();
        });

        _windowThread.SetApartmentState(ApartmentState.STA);
        _windowThread.IsBackground = true;
        _windowThread.Start();

        readyEvent.Wait();
    }

    private static void OnFormulaApplied(object? sender, string formula)
    {
        // Play chomp animation on the WPF thread (non-blocking, DispatcherTimer-based).
        // The overlay is non-focusable so it never steals focus from Excel.
        var overlay = PlayChompAnimation();

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                var worksheet = _targetWorksheet;
                var address = _targetAddress;
                if (_app == null || worksheet == null || address == null)
                {
                    return;
                }

                if (string.IsNullOrEmpty(formula))
                {
                    return;
                }

                var cell = worksheet!.Range[address];

                if (BacktickExtractor.IsBacktickFormula(formula))
                {
                    // Write as quote-prefixed text to trigger the SheetChange pipeline
                    cell.Value = "'" + formula;
                    cell.WrapText = false;
                }
                else if (formula.StartsWith('='))
                {
                    // Regular formula — write directly via Formula2 for dynamic array support
                    cell.Formula2 = formula;
                }
                else
                {
                    // Plain text/values
                    cell.Value = formula;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OnFormulaApplied error: {ex.Message}");
            }
            finally
            {
                // Signal animation to fade out now that processing is complete
                if (overlay != null)
                {
                    _windowDispatcher?.BeginInvoke(() => overlay.BeginFadeOut(2000));
                }
            }
        });
    }

    private static AnimationOverlay? PlayChompAnimation()
    {
        try
        {
            var frames = ChompAnimation.BuildFrames();
            var overlay = new AnimationOverlay(frames, 150)
            {
                OneShot = true, // Position centered on the target cell
                // WPF uses DIPs; on the animation thread we already have per-monitor DPI awareness,
                // so screen pixels from PointsToScreenPixels map 1:1 to WPF units on that monitor.
                Left = _targetCellScreenLeft - 30,
                Top = _targetCellScreenTop - 80
            };

            overlay.PlayOnce();

            return overlay;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"PlayChompAnimation error: {ex.Message}");
            return null;
        }
    }
}
