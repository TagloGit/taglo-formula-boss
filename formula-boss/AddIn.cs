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

            // Initialize the range resolver delegate for cross-sheet object model access.
            // Must be done here (host assembly context) so the lambda can call XlCall directly.
            RuntimeHelpers.ResolveRangeDelegate = rangeRef =>
            {
                var address = (string)XlCall.Excel(XlCall.xlfReftext, rangeRef, true);
                Debug.WriteLine($"ResolveRangeDelegate: {address}");
                dynamic app = ExcelDnaUtil.Application;
                return app.Range[address];
            };

            // Initialize RuntimeBridge delegates for wrapper types
            Runtime.RuntimeBridge.GetCell = (sheetName, row, col) =>
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic sheet = app.Sheets[sheetName];
                dynamic cell = sheet.Cells[row, col];
                try
                {
                    return new Runtime.Cell
                    {
                        Value = cell.Value2,
                        Formula = cell.Formula,
                        Format = cell.NumberFormat,
                        Address = cell.Address,
                        Row = cell.Row,
                        Col = cell.Column,
                        Interior = new Runtime.Interior
                        {
                            ColorIndex = cell.Interior.ColorIndex is double ci ? (int)ci : 0,
                            Color = cell.Interior.Color is double c ? (int)c : 0
                        },
                        Font = new Runtime.CellFont
                        {
                            Bold = cell.Font.Bold is true,
                            Italic = cell.Font.Italic is true,
                            Size = cell.Font.Size is double s ? s : 11,
                            Name = cell.Font.Name?.ToString() ?? "",
                            Color = cell.Font.Color is double fc ? (int)fc : 0
                        }
                    };
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }
            };

            Runtime.RuntimeBridge.GetHeaders = rangeRef =>
            {
                try
                {
                    var address = (string)XlCall.Excel(XlCall.xlfReftext, rangeRef, true);
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic range = app.Range[address];
                    try
                    {
                        dynamic listObject = range.ListObject;
                        if (listObject == null)
                        {
                            return null;
                        }

                        dynamic headerRow = listObject.HeaderRowRange;
                        if (headerRow == null)
                        {
                            return null;
                        }

                        var cols = (int)headerRow.Columns.Count;
                        var headers = new string[cols];
                        for (var i = 1; i <= cols; i++)
                        {
                            headers[i - 1] = headerRow.Cells[1, i].Value2?.ToString() ?? "";
                        }

                        return headers;
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetHeaders failed: {ex.Message}");
                    return null;
                }
            };

            Runtime.RuntimeBridge.GetOrigin = rangeRef =>
            {
                try
                {
                    var address = (string)XlCall.Excel(XlCall.xlfReftext, rangeRef, true);
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic range = app.Range[address];
                    try
                    {
                        var sheetName = (string)range.Worksheet.Name;
                        var topRow = (int)range.Row;
                        var leftCol = (int)range.Column;
                        return new Runtime.RangeOrigin(sheetName, topRow, leftCol);
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetOrigin failed: {ex.Message}");
                    return null;
                }
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
