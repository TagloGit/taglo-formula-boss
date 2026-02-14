namespace FormulaBoss.UI;

/// <summary>
///     Converts Excel cell coordinates (in points) to physical screen pixel positions.
///     Handles multi-monitor DPI differences by detecting the correct monitor via
///     MonitorFromPoint rather than relying on the system DPI.
/// </summary>
/// <remarks>
///     Excel's PointsToScreenPixelsX/Y does NOT convert points to pixels — the delta
///     from the origin equals the raw point value unchanged. The correct conversion is:
///     <c>physicalPx = origin + cellPts * (monitorDPI / 72) * (zoom / 100)</c>
///     where origin = PointsToScreenPixelsX/Y(0) in physical pixels.
/// </remarks>
public static class CellPositioner
{
    /// <summary>
    ///     Gets the physical pixel screen position of a cell's top-left corner.
    ///     Must be called on the Excel thread (needs COM access to ActiveWindow).
    /// </summary>
    /// <param name="excelWindow">The ActiveWindow COM object (provides PointsToScreenPixels and Zoom).</param>
    /// <param name="cellLeftPts">Cell.Left in points.</param>
    /// <param name="cellTopPts">Cell.Top in points.</param>
    /// <returns>Physical pixel coordinates suitable for SetWindowPos from a PER_MONITOR_AWARE_V2 thread.</returns>
    public static (int X, int Y) GetCellScreenPosition(dynamic excelWindow, double cellLeftPts, double cellTopPts)
    {
        var zoom = (int)excelWindow.Zoom;

        // Document origin in physical pixels (same value regardless of thread DPI context)
        var originX = (int)(double)excelWindow.PointsToScreenPixelsX(0);
        var originY = (int)(double)excelWindow.PointsToScreenPixelsY(0);

        // Temporarily switch to PER_MONITOR_AWARE_V2 so MonitorFromPoint interprets
        // our physical-pixel origin correctly and returns the right monitor
        var prevCtx = NativeMethods.SetThreadDpiAwarenessContext(
            NativeMethods.DpiAwarenessContextPerMonitorAwareV2);
        uint dpiX, dpiY;
        try
        {
            var originPt = new NativeMethods.Point { X = originX, Y = originY };
            var monitor = NativeMethods.MonitorFromPoint(originPt, NativeMethods.MonitorDefaultToNearest);
            _ = NativeMethods.GetDpiForMonitor(monitor, 0, out dpiX, out dpiY);
        }
        finally
        {
            NativeMethods.SetThreadDpiAwarenessContext(prevCtx);
        }

        var x = originX + (int)(cellLeftPts * (dpiX / 72.0) * (zoom / 100.0));
        var y = originY + (int)(cellTopPts * (dpiY / 72.0) * (zoom / 100.0));

        return (x, y);
    }

    /// <summary>
    ///     Positions a window at the specified physical pixel coordinates using SetWindowPos.
    ///     Must be called from a PER_MONITOR_AWARE_V2 thread for correct placement.
    /// </summary>
    public static void PlaceWindow(IntPtr hwnd, int x, int y)
    {
        NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0,
            NativeMethods.SwpNoSize | NativeMethods.SwpNoZOrder | NativeMethods.SwpNoActivate);
    }
}
