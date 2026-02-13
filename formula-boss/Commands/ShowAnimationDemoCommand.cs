using System.Diagnostics;
using System.Windows.Interop;
using System.Windows.Threading;

using ExcelDna.Integration;

using FormulaBoss.UI;
using FormulaBoss.UI.Animation;

namespace FormulaBoss.Commands;

public static class ShowAnimationDemoCommand
{
    [ExcelCommand(MenuName = "Formula Boss", MenuText = "Animation Demo")]
    public static void ShowAnimationDemo()
    {
        try
        {
            var app = ExcelDnaUtil.Application as dynamic;
            var excelHwnd = new IntPtr((int)app!.Hwnd);

            var thread = new Thread(() =>
            {
                NativeMethods.SetThreadDpiAwarenessContext(
                    NativeMethods.DpiAwarenessContextPerMonitorAwareV2);

                var frames = ChompAnimation.BuildFrames();
                var overlay = new AnimationOverlay(frames);

                // Position centered on Excel once loaded
                overlay.Loaded += (_, _) =>
                {
                    var hwnd = new WindowInteropHelper(overlay).Handle;
                    WindowPositioner.CenterOnExcel(excelHwnd, hwnd);
                };

                overlay.PlayLoop();
                Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ShowAnimationDemo error: {ex.Message}");
        }
    }
}
