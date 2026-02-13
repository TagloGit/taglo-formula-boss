using System.Diagnostics;
using System.Windows.Interop;
using System.Windows.Threading;

using ExcelDna.Integration;

using FormulaBoss.UI;

namespace FormulaBoss.Commands;

/// <summary>
///     Command handler for showing the floating formula editor.
/// </summary>
public static class ShowFloatingEditorCommand
{
    private static FloatingEditorWindow? _window;
    private static Dispatcher? _windowDispatcher;
    private static Thread? _windowThread;
    private static dynamic? _app;

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
        if (_windowDispatcher != null)
        {
            _windowDispatcher.InvokeShutdown();
            _windowDispatcher = null;
        }

        _windowThread = null;
        _window = null;
        _app = null;
    }

    /// <summary>
    ///     Toggles the floating formula editor visibility.
    ///     If hidden, shows it near the active cell.
    ///     If visible, hides it.
    /// </summary>
    [ExcelCommand(MenuName = "Formula Boss", MenuText = "Open Editor")]
    public static void ShowFloatingEditor()
    {
        try
        {
            _app ??= ExcelDnaUtil.Application;

            // Get cell info and Excel HWND while we're on the Excel thread
            var cell = _app.ActiveCell;
            var formula = cell.Formula2 as string ?? cell.Formula as string ?? "";
            var excelHwnd = new IntPtr(_app.Hwnd);

            // Create window thread if needed
            if (_windowThread == null || !_windowThread.IsAlive)
            {
                var readyEvent = new ManualResetEventSlim(false);

                _windowThread = new Thread(() =>
                {
                    // Set per-monitor v2 DPI awareness for this thread so the
                    // window renders correctly on any monitor regardless of
                    // the host process (Excel) DPI awareness mode.
                    NativeMethods.SetThreadDpiAwarenessContext(
                        NativeMethods.DpiAwarenessContextPerMonitorAwareV2);

                    _window = new FloatingEditorWindow();
                    _window.FormulaApplied += OnFormulaApplied;
                    _windowDispatcher = Dispatcher.CurrentDispatcher;
                    readyEvent.Set();

                    // Run the WPF message loop
                    Dispatcher.Run();
                });

                _windowThread.SetApartmentState(ApartmentState.STA);
                _windowThread.IsBackground = true;
                _windowThread.Start();

                // Wait for window to be created
                readyEvent.Wait();
            }

            // Dispatch to window thread
            _windowDispatcher?.Invoke(() =>
            {
                if (_window == null)
                {
                    return;
                }

                if (_window.IsVisible)
                {
                    _window.Hide();
                }
                else
                {
                    _window.FormulaText = formula;

                    // Ensure HWND exists without showing, then position, then show
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

    private static void OnFormulaApplied(object? sender, string formula)
    {
        // This runs on the WPF thread, need to dispatch to Excel thread
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                if (_app == null)
                {
                    return;
                }

                var cell = _app.ActiveCell;
                if (!string.IsNullOrEmpty(formula))
                {
                    cell.Formula2 = formula;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OnFormulaApplied error: {ex.Message}");
            }
        });
    }
}
