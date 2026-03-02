using System.Diagnostics;
using System.Runtime.InteropServices;

using Xunit;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace FormulaBoss.AddinTests;

/// <summary>
///     Manages the lifecycle of an Excel instance with the Formula Boss add-in loaded.
///     Launches a hidden Excel, registers the XLL, and cleans up on dispose.
///     Use as a collection fixture so all tests in the collection share one Excel instance.
/// </summary>
public sealed class ExcelAddinFixture : IDisposable
{
    private readonly dynamic _app;
    private readonly dynamic _workbook;
    private readonly int _excelPid;
    private bool _disposed;

    public ExcelAddinFixture()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel is not installed or not registered.");

        // Snapshot Excel PIDs before launch so we can identify the one we spawned
        var pidsBefore = new HashSet<int>(
            Process.GetProcessesByName("EXCEL").Select(p => p.Id));

        _app = Activator.CreateInstance(excelType)
               ?? throw new InvalidOperationException("Failed to create Excel.Application instance.");

        _app.Visible = false;
        _app.DisplayAlerts = false;

        // Identify the new Excel process by diffing PIDs
        _excelPid = FindNewExcelPid(pidsBefore);

        // Load the Formula Boss XLL
        var xllPath = FindXllPath();
        bool registered = _app.RegisterXLL(xllPath);
        if (!registered)
        {
            _app.Quit();
            Marshal.ReleaseComObject(_app);
            throw new InvalidOperationException($"Failed to register XLL: {xllPath}");
        }

        // Create a workbook for tests
        _workbook = _app.Workbooks.Add();

        // Wait for AutoOpen to initialize the interceptor — poll rather than fixed sleep
        WaitForAddinReady();
    }

    /// <summary>
    ///     The Excel Application COM object.
    /// </summary>
    public dynamic Application => _app;

    /// <summary>
    ///     The test workbook.
    /// </summary>
    public dynamic Workbook => _workbook;

    /// <summary>
    ///     Gets the first worksheet in the test workbook.
    /// </summary>
    public dynamic Worksheet => _workbook.Worksheets[1];

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        try
        {
            _workbook.Close(false);
            Marshal.ReleaseComObject(_workbook);
        }
        catch
        {
            // Ignore cleanup errors
        }

        try
        {
            _app.Quit();
            Marshal.ReleaseComObject(_app);
        }
        catch
        {
            // Ignore cleanup errors
        }

        // Safety net: kill our specific Excel process if Quit() didn't work
        KillSpawnedExcel();
    }

    /// <summary>
    ///     Adds a fresh worksheet for test isolation.
    ///     Each test should call this to avoid interference.
    /// </summary>
    public dynamic AddWorksheet()
    {
        return _workbook.Worksheets.Add();
    }

    /// <summary>
    ///     Polls until the add-in's SheetChange event handler is wired up,
    ///     which we detect by entering a backtick formula and checking if it gets rewritten.
    /// </summary>
    private void WaitForAddinReady(int timeoutMs = 10000, int pollIntervalMs = 250)
    {
        var ws = _workbook.Worksheets[1];
        var cell = ws.Range["ZZ1"];
        try
        {
            // Enter a minimal backtick formula
            cell.Value = "'=`ZZ2:ZZ2.Sum()`";

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var formula = cell.Formula2 as string;
                    // If the interceptor has rewritten it (no more backticks), the add-in is ready
                    if (formula != null && formula.StartsWith('=') && !formula.Contains('`'))
                    {
                        // Clean up the probe cell
                        cell.ClearContents();
                        return;
                    }
                }
                catch
                {
                    // COM might fail during init
                }

                Thread.Sleep(pollIntervalMs);
            }

            // Timed out — clean up and proceed anyway (tests may still work with slight delay)
            try
            {
                cell.ClearContents();
            }
            catch
            {
                // Ignore
            }
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(ws);
        }
    }

    /// <summary>
    ///     Locates the 64-bit XLL in the build output directory.
    /// </summary>
    private static string FindXllPath()
    {
        // Walk up from the test assembly output to find the formula-boss build output
        var testDir = AppContext.BaseDirectory;
        var repoRoot = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", ".."));

        // Try Debug first, then Release
        foreach (var config in new[] { "Debug", "Release" })
        {
            var xllPath = Path.Combine(repoRoot, "formula-boss", "bin", config,
                "net6.0-windows", "formula-boss64.xll");
            if (File.Exists(xllPath))
            {
                return xllPath;
            }
        }

        throw new FileNotFoundException(
            $"Could not find formula-boss64.xll. Build the formula-boss project first. Searched from: {repoRoot}");
    }

    /// <summary>
    ///     Finds the Excel process we spawned by comparing PIDs before and after launch.
    /// </summary>
    private static int FindNewExcelPid(HashSet<int> pidsBefore, int timeoutMs = 5000)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            foreach (var proc in Process.GetProcessesByName("EXCEL"))
            {
                if (!pidsBefore.Contains(proc.Id))
                {
                    return proc.Id;
                }
            }

            Thread.Sleep(100);
        }

        return -1; // Couldn't identify — fall back to no-kill behavior
    }

    /// <summary>
    ///     Kills only the Excel process we spawned, identified by PID.
    /// </summary>
    private void KillSpawnedExcel()
    {
        if (_excelPid <= 0)
        {
            return;
        }

        try
        {
            var proc = Process.GetProcessById(_excelPid);
            if (proc is { HasExited: false, ProcessName: "EXCEL" })
            {
                proc.Kill();
            }
        }
        catch
        {
            // Process already exited or access denied — fine
        }
    }
}

/// <summary>
///     xUnit collection definition that shares a single Excel instance across all test classes.
/// </summary>
[CollectionDefinition("Excel Addin")]
public class ExcelAddinCollection : ICollectionFixture<ExcelAddinFixture>
{
}
