using System.Diagnostics;

using ExcelDna.Integration;

using FormulaBoss.Commands;
using FormulaBoss.Compilation;
using FormulaBoss.Interception;

namespace FormulaBoss;

/// <summary>
/// Excel add-in entry point. Handles registration of static and dynamic UDFs.
/// </summary>
public sealed class AddIn : IExcelAddIn, IDisposable
{
    private DynamicCompiler? _compiler;
    private FormulaPipeline? _pipeline;
    private FormulaInterceptor? _interceptor;
    private bool _disposed;

    public void AutoOpen()
    {
        try
        {
            // Initialize the range resolver delegate for cross-sheet object model access.
            // Must be done here (host assembly context) so the lambda can call XlCall directly.
            RuntimeHelpers.ResolveRangeDelegate = rangeRef =>
            {
                var address = (string)XlCall.Excel(XlCall.xlfReftext, rangeRef, true);
                Debug.WriteLine($"ResolveRangeDelegate: {address}");
                dynamic app = ExcelDnaUtil.Application;
                return app.Range[address];
            };

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

    public void AutoClose()
    {
        Dispose();
    }

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
            // Ignore errors during cleanup
        }

        ShowFloatingEditorCommand.Cleanup();

        _interceptor?.Dispose();
        _interceptor = null;
        _pipeline = null;
        _compiler = null;
        _disposed = true;
    }

}
