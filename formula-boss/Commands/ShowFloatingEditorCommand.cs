using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Threading;

using ExcelDna.Integration;

using FormulaBoss.Interception;
using FormulaBoss.UI;
using FormulaBoss.UI.Animation;

using Taglo.Excel.Common;

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

    // Cell screen position captured at open time for animation placement (physical pixels)
    private static int _targetCellScreenLeft;
    private static int _targetCellScreenTop;
    private static double _targetMonitorScale = 1.0;

    // True after the editor has been positioned at least once — subsequent opens
    // reuse whatever position the user left the window in (session-only memory).
    private static bool _hasBeenPositioned;

    /// <summary>
    ///     Initializes the command with the Excel application reference.
    ///     Must be called during add-in initialization.
    /// </summary>
    public static void Initialize(dynamic app) => _app = app;

    /// <summary>
    ///     Cleans up the floating editor window and releases all COM references.
    ///     Must be called during add-in shutdown.
    /// </summary>
    public static void Cleanup()
    {
        // 1. Release stored COM references BEFORE shutting down threads
        ReleaseTargetWorksheet();

        // 2. Shut down the WPF dispatcher
        _windowDispatcher?.InvokeShutdown();

        // 3. Wait for the WPF thread to actually finish (InvokeShutdown is async)
        if (_windowThread is { IsAlive: true })
        {
            _windowThread.Join(TimeSpan.FromSeconds(2));
        }

        // 4. Null out all references
        _windowDispatcher = null;
        _windowThread = null;
        _window = null;
        _app = null;
        _targetAddress = null;
        _hasBeenPositioned = false;
    }

    /// <summary>
    ///     Opens the floating formula editor with context-aware content based on cell state.
    ///     If editor is already visible, hides it (toggle).
    /// </summary>
    [ExcelCommand]
    public static void ShowFloatingEditor()
    {
        dynamic? cell = null;
        dynamic? worksheet = null;

        try
        {
            _app ??= ExcelDnaUtil.Application;

            cell = _app.ActiveCell;
            var currentAddress = cell.Address as string;
            var editorContent = DetectEditorContent(cell);
            var excelHwnd = new IntPtr(_app.Hwnd);
            var metadata = WorkbookMetadata.CaptureFromExcel(_app);

            // Capture worksheet on the Excel thread (not inside the WPF dispatcher)
            // to avoid cross-apartment COM proxy complications
            worksheet = cell.Worksheet;
            var sheetName = worksheet.Name as string;

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
                        ReleaseTargetWorksheet();
                        _targetWorksheet = worksheet;
                        worksheet = null; // Ownership transferred
                        _targetAddress = currentAddress;
                        _window.UpdateTitle(sheetName, currentAddress);
                        _window.Metadata = metadata;
                        _window.FormulaText = editorContent;
                    }
                }
                else
                {
                    // Capture target cell at open time
                    ReleaseTargetWorksheet();
                    _targetWorksheet = worksheet;
                    worksheet = null; // Ownership transferred
                    _targetAddress = currentAddress;
                    _window.UpdateTitle(sheetName, currentAddress);
                    _window.Metadata = metadata;
                    _window.FormulaText = editorContent;

                    if (!_hasBeenPositioned)
                    {
                        var wpfHwnd = new WindowInteropHelper(_window).EnsureHandle();
                        WindowPositioner.CenterOnExcel(excelHwnd, wpfHwnd);
                        _hasBeenPositioned = true;
                    }

                    _window.Show();
                    _window.Activate();
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Error("ShowFloatingEditor", ex);
        }
        finally
        {
            // Release transient COM objects that weren't transferred to _targetWorksheet
            ReleaseCom(worksheet);
            ReleaseCom(cell);
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
            var text = cell.Value2 as string ?? "";
            return FormatLetIfEnabled(text);
        }

        // Processed FB LET formula: contains _src_ variables from pipeline
        if (LetFormulaReconstructor.IsProcessedFormulaBossLet(formula))
        {
            if (LetFormulaReconstructor.TryReconstruct(formula, out var editableFormula) &&
                editableFormula != null)
            {
                // TryReconstruct adds a leading ' for text storage — strip it for the editor
                var text = editableFormula.StartsWith('\'') ? editableFormula[1..] : editableFormula;
                return FormatLetIfEnabled(text);
            }
        }

        // Processed basic FB / non-FB formula: format LET if applicable
        return FormatLetIfEnabled(formula);
    }

    /// <summary>
    ///     Formats a formula using LetFormulaFormatter if AutoFormatLet is enabled
    ///     and the formula is a LET formula.
    /// </summary>
    private static string FormatLetIfEnabled(string formula)
    {
        var settings = EditorSettings.Load();
        if (!settings.AutoFormatLet)
        {
            return formula;
        }

        return LetFormulaFormatter.Format(
            formula,
            settings.IndentSize,
            settings.NestedLetDepth,
            settings.MaxLineLength);
    }

    /// <summary>
    ///     Captures the target cell's screen position in physical pixels.
    /// </summary>
    private static void CaptureTargetCellPosition(dynamic cell)
    {
        dynamic? window = null;
        try
        {
            window = _app!.ActiveWindow;
            double cellLeft = cell.Left;
            double cellTop = cell.Top;
            var pos = ((int X, int Y, double Scale))CellPositioner.GetCellScreenPosition(window, cellLeft, cellTop);
            _targetCellScreenLeft = pos.X;
            _targetCellScreenTop = pos.Y;
            _targetMonitorScale = pos.Scale;
        }
        catch (Exception ex)
        {
            Logger.Error("CaptureTargetCellPosition", ex);
            _targetCellScreenLeft = 0;
            _targetCellScreenTop = 0;
        }
        finally
        {
            ReleaseCom(window);
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

            // Catch unhandled exceptions on the WPF thread so they don't crash Excel
            _windowDispatcher.UnhandledException += (_, e) =>
            {
                Logger.Error("WPF dispatcher", e.Exception);
                e.Handled = true;
            };

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
        // Play animation on the WPF thread (non-blocking, DispatcherTimer-based).
        // The overlay is non-focusable so it never steals focus from Excel.
        var overlay = PlayAnimation();

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            dynamic? cell = null;
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

                cell = worksheet!.Range[address];

                if (BacktickExtractor.IsBacktickFormula(formula))
                {
                    // Write as quote-prefixed text to trigger the SheetChange pipeline
                    cell.Value = "'" + formula;
                    cell.WrapText = false;
                }
                else if (formula.StartsWith('='))
                {
                    // Format LET formulas before writing to the cell
                    var formatted = FormatLetIfEnabled(formula);

                    // Regular formula — write directly via Formula2 for dynamic array support
                    cell.Formula2 = formatted;
                }
                else
                {
                    // Plain text/values
                    cell.Value = formula;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("OnFormulaApplied", ex);
            }
            finally
            {
                ReleaseCom(cell);

                // Signal animation to fade out now that processing is complete
                if (overlay != null)
                {
                    _windowDispatcher?.BeginInvoke(() => overlay.BeginFadeOut(1500));
                }
            }
        });
    }

    private static AnimationOverlay? PlayAnimation()
    {
        try
        {
            var settings = EditorSettings.Load();
            if (settings.AnimationStyle == AnimationStyle.None)
            {
                return null;
            }

            var frames = settings.AnimationStyle switch
            {
                AnimationStyle.Roar => RoarAnimation.BuildFrames(),
                AnimationStyle.Shuffle => ShuffleAnimation.BuildFrames(),
                _ => ChompAnimation.BuildFrames()
            };

            var overlay = new AnimationOverlay(frames, 100) { OneShot = true };

            // Position the native window BEFORE Show() so the first rendered frame
            // is already at the correct location. EnsureHandle() creates the HWND
            // without making the window visible.
            // Offset in logical pixels (at 96 DPI), scaled to physical pixels for the monitor
            var offsetX = (int)(-20 * _targetMonitorScale);
            var offsetY = (int)(-40 * _targetMonitorScale);

            var hwnd = new WindowInteropHelper(overlay).EnsureHandle();
            CellPositioner.PlaceWindow(hwnd,
                _targetCellScreenLeft + offsetX,
                _targetCellScreenTop + offsetY);

            overlay.PlayOnce();

            return overlay;
        }
        catch (Exception ex)
        {
            Logger.Error("PlayAnimation", ex);
            return null;
        }
    }

    /// <summary>
    ///     Opens the settings dialog on the WPF thread.
    /// </summary>
    public static void ShowSettings()
    {
        try
        {
            _app ??= ExcelDnaUtil.Application;
            var excelHwnd = new IntPtr(_app.Hwnd);

            EnsureWindowThread();

            _windowDispatcher?.Invoke(() =>
            {
                var settings = EditorSettings.Load();
                var dialog = new SettingsDialog(settings);

                if (_window is { IsVisible: true })
                {
                    // Editor is open — position on top of it
                    dialog.Owner = _window;
                }
                else
                {
                    // Editor not open — center on Excel window
                    var dialogHwnd = new WindowInteropHelper(dialog).EnsureHandle();
                    WindowPositioner.CenterOnExcel(excelHwnd, dialogHwnd);
                }

                if (dialog.ShowDialog() == true)
                {
                    settings.AnimationStyle = dialog.SelectedAnimation;
                    settings.IndentSize = dialog.SelectedIndentSize;
                    settings.WordWrap = dialog.SelectedWordWrap;
                    settings.AutoFormatLet = dialog.SelectedAutoFormatLet;
                    settings.NestedLetDepth = dialog.SelectedNestedLetDepth;
                    settings.MaxLineLength = dialog.SelectedMaxLineLength;
                    settings.Save();

                    if (_window is { IsVisible: true })
                    {
                        _window.ApplySettings(settings);
                    }
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Error("ShowSettings", ex);
        }
    }

    /// <summary>
    ///     Opens the About dialog on the WPF thread.
    /// </summary>
    public static void ShowAbout()
    {
        try
        {
            _app ??= ExcelDnaUtil.Application;
            var excelHwnd = new IntPtr(_app.Hwnd);

            EnsureWindowThread();

            _windowDispatcher?.Invoke(() =>
            {
                var dialog = new AboutDialog();

                if (_window is { IsVisible: true })
                {
                    dialog.Owner = _window;
                }
                else
                {
                    var dialogHwnd = new WindowInteropHelper(dialog).EnsureHandle();
                    WindowPositioner.CenterOnExcel(excelHwnd, dialogHwnd);
                }

                dialog.ShowDialog();
            });
        }
        catch (Exception ex)
        {
            Logger.Error("ShowAbout", ex);
        }
    }

    /// <summary>
    ///     Releases the stored target worksheet COM reference.
    /// </summary>
    private static void ReleaseTargetWorksheet()
    {
        if (_targetWorksheet != null)
        {
            try
            {
                Marshal.ReleaseComObject(_targetWorksheet);
            }
            catch
            {
                // Ignore — object may already be released or invalid
            }

            _targetWorksheet = null;
        }
    }

    /// <summary>
    ///     Safely releases a COM object if non-null.
    /// </summary>
    private static void ReleaseCom(object? comObject)
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Ignore — object may not be a COM object or already released
            }
        }
    }
}
