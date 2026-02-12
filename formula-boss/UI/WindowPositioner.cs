namespace FormulaBoss.UI;

public static class WindowPositioner
{
    /// <summary>
    ///     Gets the screen position to center a window over the Excel window.
    ///     Thread-safe: uses only Win32 calls, no COM interop.
    /// </summary>
    /// <summary>
    ///     Centers the WPF window over the Excel window using Win32 SetWindowPos.
    ///     Works entirely in physical pixels to avoid DPI context mismatches
    ///     between monitors with different scaling.
    /// </summary>
    public static void CenterOnExcel(IntPtr excelHwnd, IntPtr wpfHwnd)
    {
        if (!NativeMethods.GetWindowRect(excelHwnd, out var excelRect))
        {
            return;
        }

        if (!NativeMethods.GetWindowRect(wpfHwnd, out var wpfRect))
        {
            return;
        }

        var wpfWidth = wpfRect.Width;
        var wpfHeight = wpfRect.Height;

        var left = excelRect.Left + ((excelRect.Width - wpfWidth) / 2);
        var top = excelRect.Top + ((excelRect.Height - wpfHeight) / 2);

        // Constrain to the work area of Excel's monitor
        var monitor = NativeMethods.MonitorFromWindow(excelHwnd, NativeMethods.MonitorDefaultToNearest);
        var monitorInfo = NativeMethods.Monitorinfo.Create();
        NativeMethods.GetMonitorInfo(monitor, ref monitorInfo);
        var workArea = monitorInfo.rcWork;

        left = Math.Max(workArea.Left, Math.Min(left, workArea.Right - wpfWidth));
        top = Math.Max(workArea.Top, Math.Min(top, workArea.Bottom - wpfHeight));

        NativeMethods.SetWindowPos(wpfHwnd, IntPtr.Zero, left, top, 0, 0,
            NativeMethods.SwpNoSize | NativeMethods.SwpNoZOrder | NativeMethods.SwpNoActivate);
    }
}
