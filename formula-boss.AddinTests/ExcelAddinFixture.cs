using System.Diagnostics;
using System.Runtime.InteropServices;

using Xunit;
using Xunit.Abstractions;

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
    private bool _disposed;

    public ExcelAddinFixture()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("Excel is not installed or not registered.");

        _app = Activator.CreateInstance(excelType)
               ?? throw new InvalidOperationException("Failed to create Excel.Application instance.");

        _app.Visible = false;
        _app.DisplayAlerts = false;

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

        // Give AutoOpen time to initialize the interceptor and pipeline
        Thread.Sleep(3000);
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

        // Kill any orphaned Excel processes we might have spawned
        KillOrphanedExcel();
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

    private static void KillOrphanedExcel()
    {
        // Only kill Excel processes that were started very recently and are hidden
        // This is a safety net — normally Quit() handles it
        try
        {
            foreach (var proc in Process.GetProcessesByName("EXCEL"))
            {
                if (proc.MainWindowHandle == IntPtr.Zero) // Hidden instance
                {
                    proc.Kill();
                }
            }
        }
        catch
        {
            // Best effort
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
