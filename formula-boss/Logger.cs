using System.Diagnostics;

namespace FormulaBoss;

/// <summary>
///     Simple file logger for crash diagnostics. Writes to %LOCALAPPDATA%\FormulaBoss\logs\formulaboss.log.
///     All I/O failures are silently ignored — the logger must never itself cause a crash.
/// </summary>
internal static class Logger
{
    private const long MaxFileSize = 1_048_576; // 1 MB

    private static readonly object Lock = new();
    private static string? _logFilePath;

    /// <summary>
    ///     Initializes the logger. Truncates the log file if it exceeds 1 MB.
    ///     Call once during AddIn.AutoOpen.
    /// </summary>
    public static void Initialize()
    {
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "FormulaBoss",
                "logs");
            Directory.CreateDirectory(dir);
            _logFilePath = Path.Combine(dir, "formulaboss.log");

            // Truncate if over 1 MB
            if (File.Exists(_logFilePath) && new FileInfo(_logFilePath).Length > MaxFileSize)
            {
                File.WriteAllText(_logFilePath, "");
            }
        }
        catch
        {
            // Silently ignore — logging is best-effort
        }
    }

    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Error(string source, Exception ex)
    {
        Write("ERROR", $"{source}: {ex}");
    }

    private static void Write(string level, string message)
    {
        // Always write to Debug output as before
        Debug.WriteLine($"[{level}] {message}");

        if (_logFilePath == null)
        {
            return;
        }

        try
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}{Environment.NewLine}";
            lock (Lock)
            {
                File.AppendAllText(_logFilePath, line);
            }
        }
        catch
        {
            // Silently ignore — logging is best-effort
        }
    }
}
