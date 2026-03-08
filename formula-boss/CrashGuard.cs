using System.Diagnostics;
using System.IO;

namespace FormulaBoss;

/// <summary>
///     Global exception guards and logging to prevent the add-in from crashing Excel.
///     All unhandled exceptions are caught, logged, and swallowed.
/// </summary>
internal static class CrashGuard
{
    private static readonly object Lock = new();
    private static string? _logPath;

    /// <summary>
    ///     Absolute path to the current log file.
    /// </summary>
    public static string LogPath => _logPath ??= InitLogPath();

    /// <summary>
    ///     Installs global exception handlers that prevent unhandled exceptions
    ///     from propagating to Excel and crashing the host process.
    ///     Must be called once, as early as possible in <see cref="AddIn.AutoOpen" />.
    /// </summary>
    public static void InstallGlobalHandlers()
    {
        // Catch exceptions on threads that have no try/catch (thread pool, manual threads).
        // IsTerminating is true for .NET Framework but always false for .NET Core/5+;
        // either way we log and move on.
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            Log("AppDomain.UnhandledException", args.ExceptionObject as Exception);
        };

        // Catch unobserved Task exceptions (fire-and-forget async, un-awaited tasks).
        // SetObserved() prevents the runtime from tearing down the process.
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            Log("TaskScheduler.UnobservedTaskException", args.Exception);
            args.SetObserved();
        };
    }

    /// <summary>
    ///     Logs an error with context to the log file and Debug output.
    /// </summary>
    public static void Log(string context, Exception? ex)
    {
        var message = ex != null
            ? $"{context}: {ex}"
            : $"{context}: (no exception object)";

        Debug.WriteLine($"[CrashGuard] {message}");

        try
        {
            lock (Lock)
            {
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Last resort — can't even write to log. Nothing more we can do.
        }
    }

    /// <summary>
    ///     Logs a plain message (no exception) to the log file and Debug output.
    /// </summary>
    public static void Log(string message)
    {
        Debug.WriteLine($"[CrashGuard] {message}");

        try
        {
            lock (Lock)
            {
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Last resort
        }
    }

    private static string InitLogPath()
    {
        // Place log file in %LOCALAPPDATA%/FormulaBoss/ so it survives add-in updates
        // and is accessible even when the XLL directory is read-only.
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FormulaBoss");

        try
        {
            Directory.CreateDirectory(dir);
        }
        catch
        {
            // Fall back to temp directory if we can't create the folder
            dir = Path.GetTempPath();
        }

        return Path.Combine(dir, "formula-boss.log");
    }
}
