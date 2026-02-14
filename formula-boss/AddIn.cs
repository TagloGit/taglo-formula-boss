using System.Diagnostics;

using ExcelDna.Integration;

using FormulaBoss.Commands;
using FormulaBoss.Compilation;
using FormulaBoss.Interception;

namespace FormulaBoss;

/// <summary>
///     Excel add-in entry point. Handles registration of static and dynamic UDFs.
/// </summary>
public sealed class AddIn : IExcelAddIn, IDisposable
{
    private static AddIn? _instance;

    private DynamicCompiler? _compiler;
    private bool _disposed;
    private FormulaInterceptor? _interceptor;
    private FormulaPipeline? _pipeline;

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        // Unregister keyboard shortcuts
        try
        {
            XlCall.Excel(XlCall.xlcOnKey, "^+`");
        }
        catch
        {
            // Ignore errors during cleanup — Excel may already be shutting down
        }

        ShowFloatingEditorCommand.Cleanup();

        _interceptor?.Dispose();
        _interceptor = null;
        _pipeline = null;
        _compiler = null;
        _disposed = true;
    }

    public void AutoOpen()
    {
        _instance = this;

        try
        {
            // AutoClose is NOT called when Excel shuts down — only when the
            // add-in is explicitly removed via the Add-Ins dialog.
            // Register a COM add-in to get OnBeginShutdown notification,
            // which is the ExcelDNA-recommended way to detect Excel closing.
            ExcelComAddInHelper.LoadComAddIn(new ShutdownMonitor());

            // Defer event hookup until Excel is fully initialized
            // ExcelAsyncUtil.QueueAsMacro ensures we run after AutoOpen completes
            ExcelAsyncUtil.QueueAsMacro(InitializeInterception);

            Debug.WriteLine("Formula Boss add-in loaded successfully");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"AddIn.AutoOpen error: {ex}");
        }
    }

    public void AutoClose() => Dispose();

    private void InitializeInterception()
    {
        try
        {
            // Initialize the dynamic compilation infrastructure
            _compiler = new DynamicCompiler();
            _pipeline = new FormulaPipeline(_compiler);
            _interceptor = new FormulaInterceptor(_pipeline);

            // Start listening for worksheet changes
            _interceptor.Start();

            // Register keyboard shortcut: Ctrl+Shift+` to open floating editor
            // ^+` = Ctrl+Shift+` (^ = Ctrl, + = Shift)
            XlCall.Excel(XlCall.xlcOnKey, "^+`", "ShowFloatingEditor");

            Debug.WriteLine("Formula Boss interception initialized");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"InitializeInterception error: {ex}");
        }
    }

    /// <summary>
    ///     COM add-in registered solely to receive Excel shutdown notification.
    /// </summary>
    private class ShutdownMonitor : ExcelComAddIn
    {
        public override void OnBeginShutdown(ref Array custom) => _instance?.Dispose();
    }
}
