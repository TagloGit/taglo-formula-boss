using System.Collections;
using System.Diagnostics;

using ExcelDna.Integration;

using FormulaBoss.Commands;
using FormulaBoss.Compilation;
using FormulaBoss.Interception;
using FormulaBoss.Runtime;

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

            // Initialize header extraction delegate for generated code
            RuntimeHelpers.GetHeadersDelegate = rangeRef =>
            {
                try
                {
                    var values = RuntimeHelpers.GetValuesFromReference(rangeRef);
                    if (values.GetLength(0) < 1)
                    {
                        return null;
                    }

                    var cols = values.GetLength(1);
                    var headers = new string[cols];
                    for (var i = 0; i < cols; i++)
                    {
                        headers[i] = values[0, i]?.ToString() ?? "";
                    }

                    return headers;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetHeadersDelegate error: {ex.Message}");
                    return null;
                }
            };

            // Initialize origin extraction delegate for generated code
            RuntimeHelpers.GetOriginDelegate = rangeRef =>
            {
                try
                {
                    if (rangeRef?.GetType().Name != "ExcelReference")
                    {
                        return null;
                    }

                    // Get sheet name from SheetId via reflection
                    var sheetIdProp = rangeRef.GetType().GetProperty("SheetId");
                    var sheetId = sheetIdProp?.GetValue(rangeRef);
                    var sheetName = sheetId != null
                        ? (string)XlCall.Excel(XlCall.xlSheetNm, rangeRef)
                        : "Sheet1";

                    // Strip the [Book]Sheet format to just the sheet name
                    var bracketEnd = sheetName.IndexOf(']');
                    if (bracketEnd >= 0)
                    {
                        sheetName = sheetName[(bracketEnd + 1)..];
                    }

                    // Get row/col from RowFirst/ColumnFirst properties
                    var rowFirstProp = rangeRef.GetType().GetProperty("RowFirst");
                    var colFirstProp = rangeRef.GetType().GetProperty("ColumnFirst");
                    var row = (int)(rowFirstProp?.GetValue(rangeRef) ?? 0) + 1; // Convert 0-based to 1-based
                    var col = (int)(colFirstProp?.GetValue(rangeRef) ?? 0) + 1;

                    return new RangeOrigin(sheetName, row, col);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetOriginDelegate error: {ex.Message}");
                    return null;
                }
            };

            // Initialize result conversion delegate for generated code
            RuntimeHelpers.ToResultDelegate = result =>
            {
                if (result == null)
                {
                    return string.Empty;
                }

                if (result is ExcelValue ev)
                {
                    return ev.ToResult();
                }

                if (result is IExcelRange range)
                {
                    return range.ToResult();
                }

                if (result is bool b)
                {
                    return b.ToResult();
                }

                if (result is int i)
                {
                    return i.ToResult();
                }

                if (result is double d)
                {
                    return d.ToResult();
                }

                if (result is string s)
                {
                    return s.ToResult();
                }

                // Handle LINQ IEnumerable<Row> results (from .Rows.Where(), etc.)
                if (result is IEnumerable<Row> rows)
                {
                    var rowList = rows.ToList();
                    if (rowList.Count == 0)
                    {
                        return string.Empty;
                    }

                    var cols = rowList[0].ColumnCount;
                    var arr = new object?[rowList.Count, cols];
                    for (var r = 0; r < rowList.Count; r++)
                    for (var c = 0; c < cols; c++)
                    {
                        arr[r, c] = rowList[r][c].Value;
                    }

                    return arr;
                }

                // Handle IEnumerable<ColumnValue> results
                if (result is IEnumerable<ColumnValue> colValues)
                {
                    var list = colValues.ToList();
                    if (list.Count == 0)
                    {
                        return string.Empty;
                    }

                    var arr = new object?[list.Count, 1];
                    for (var r = 0; r < list.Count; r++)
                    {
                        arr[r, 0] = list[r].Value;
                    }

                    return arr;
                }

                // Handle generic IEnumerable
                if (result is IEnumerable enumerable and not string and not object[,])
                {
                    var list = enumerable.Cast<object>().ToList();
                    if (list.Count == 0)
                    {
                        return string.Empty;
                    }

                    var arr = new object?[list.Count, 1];
                    for (var r = 0; r < list.Count; r++)
                    {
                        arr[r, 0] = list[r] is ColumnValue cv ? cv.Value : list[r];
                    }

                    return arr;
                }

                return result;
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
