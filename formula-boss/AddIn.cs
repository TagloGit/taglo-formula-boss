using System.Diagnostics;
using System.Runtime.InteropServices;

using ExcelDna.Integration;

using FormulaBoss.Commands;
using FormulaBoss.Compilation;
using FormulaBoss.Interception;
using FormulaBoss.Runtime;
using FormulaBoss.Transpilation;
using Taglo.Excel.Common;

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
    private bool _isExcelShutdown;
    private FormulaPipeline? _pipeline;

    /// <summary>
    ///     The active pipeline instance, used by <see cref="Commands.DebugToggleService"/> to
    ///     compile debug variants on demand.
    /// </summary>
    internal static FormulaPipeline? Pipeline => _instance?._pipeline;

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        Logger.Info("Formula Boss add-in shutting down");

        TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;

        // Unregister keyboard shortcuts — but only when the add-in is being removed
        // while Excel stays running (AutoClose). During Excel shutdown (OnBeginShutdown),
        // the C API macro context is not available and xlcOnKey blocks indefinitely.
        if (!_isExcelShutdown)
        {
            try
            {
                XlCall.Excel(XlCall.xlcOnKey, "^+`");
            }
            catch
            {
                // Ignore errors during cleanup
            }
        }

        ShowFloatingEditorCommand.Cleanup();

        _interceptor?.Dispose();
        _interceptor = null;
        _pipeline = null;
        _compiler = null;
        _disposed = true;

        Logger.Info("Formula Boss add-in shutdown complete");
    }

    public void AutoOpen()
    {
        _instance = this;

        Logger.Initialize("FormulaBoss");
        Logger.Info("Formula Boss add-in loading");

        // Catch exceptions from fire-and-forget tasks so they don't crash the process
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

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

            // Initialize header extraction delegate for generated code.
            // Accepts already-extracted object[,] values, not raw ExcelReference.
            RuntimeHelpers.GetHeadersDelegate = values =>
            {
                try
                {
                    if (values.GetLength(0) < 1)
                    {
                        return null;
                    }

                    var cols = values.GetLength(1);
                    var headers = new string[cols];
                    for (var i = 0; i < cols; i++)
                    {
                        headers[i] = values[0, i].ToString() ?? "";
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
                    if (rangeRef.GetType().Name != "ExcelReference")
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

            // Initialize cell resolver delegate for object model access (formatting, color, etc.)
            RuntimeBridge.GetCell = (sheetName, row, col) =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    var sheet = app.Sheets[sheetName];
                    var cell = sheet.Cells[row, col];
                    var interior = cell.Interior;
                    var font = cell.Font;
                    var result = new Cell
                    {
                        Value = cell.Value2,
                        Formula = cell.Formula ?? "",
                        Format = cell.NumberFormat ?? "",
                        Address = cell.Address ?? "",
                        Row = row,
                        Col = col,
                        Interior = new Interior(
                            (int)(interior.ColorIndex ?? 0),
                            (int)(double)(interior.Color ?? 0.0)),
                        Font = new CellFont(
                            (bool)(font.Bold ?? false),
                            (bool)(font.Italic ?? false),
                            (double)(font.Size ?? 11.0),
                            (string)(font.Name ?? "Calibri"),
                            (int)(double)(font.Color ?? 0.0))
                    };
                    return result;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GetCell error: {ex.Message}");
                    return new Cell { Row = row, Col = col };
                }
            };

            // Initialize result conversion delegate — delegates to shared ResultConverter.Convert()
            RuntimeHelpers.ToResultDelegate = result => ResultConverter.Convert(result);

            // Reset the debug-mode trace buffer and wire up the caller-address delegate
            // so debug-instrumented UDFs can pass the calling cell address to Tracer.Begin.
            Tracer.Reset();
            RuntimeHelpers.GetCallerAddressDelegate = () =>
            {
                try
                {
                    var caller = XlCall.Excel(XlCall.xlfCaller);
                    if (caller?.GetType().Name == "ExcelReference")
                    {
                        return (string)XlCall.Excel(XlCall.xlfReftext, caller, true);
                    }

                    return "";
                }
                catch
                {
                    return "";
                }
            };

            // Defer event hookup until Excel is fully initialized
            // ExcelAsyncUtil.QueueAsMacro ensures we run after AutoOpen completes
            ExcelAsyncUtil.QueueAsMacro(InitializeInterception);

            // Fire-and-forget update check — runs on background thread, silent on failure
            var assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
                                  ?? new Version(0, 0, 0);
            UpdateChecker.Initialize(
                "https://api.github.com/repos/TagloGit/taglo-formula-boss/releases/latest",
                $"FormulaBoss/{assemblyVersion}");
            UpdateChecker.CheckForUpdateAsync(assemblyVersion);

            Logger.Info("Formula Boss add-in loaded successfully");
        }
        catch (Exception ex)
        {
            Logger.Error("AddIn.AutoOpen", ex);
        }
    }

    public void AutoClose() => Dispose();

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        Logger.Error("Unobserved task", e.Exception);
        e.SetObserved();
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

            // Listen for workbook open events to rehydrate debug variants
            dynamic app = ExcelDnaUtil.Application;
            app.WorkbookOpen += new WorkbookOpenHandler(OnWorkbookOpen);

            // Register keyboard shortcut: Ctrl+Shift+` to open floating editor
            // ^+` = Ctrl+Shift+` (^ = Ctrl, + = Shift)
            XlCall.Excel(XlCall.xlcOnKey, "^+`", "ShowFloatingEditor");

            Logger.Info("Formula Boss interception initialized");
        }
        catch (Exception ex)
        {
            Logger.Error("InitializeInterception", ex);
        }
    }

    private delegate void WorkbookOpenHandler(dynamic workbook);

    private void OnWorkbookOpen(dynamic workbook)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                RehydrateDebugVariants(workbook);
            }
            catch (Exception ex)
            {
                Logger.Error("OnWorkbookOpen rehydration", ex);
            }
        });
    }

    /// <summary>
    ///     Scans all sheets in the workbook for cells with _DEBUG call sites and
    ///     compiles the debug variant for each, so debug mode survives file reopen.
    /// </summary>
    private void RehydrateDebugVariants(dynamic workbook)
    {
        if (_pipeline == null)
        {
            return;
        }

        var sheets = workbook.Worksheets;
        var sheetCount = (int)sheets.Count;

        for (var i = 1; i <= sheetCount; i++)
        {
            dynamic? sheet = null;
            dynamic? usedRange = null;
            try
            {
                sheet = sheets[i];
                usedRange = sheet.UsedRange;
                ScanRangeForDebugCallSites(usedRange);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Rehydration error on sheet {i}: {ex.Message}");
            }
            finally
            {
                if (usedRange != null)
                {
                    Marshal.ReleaseComObject(usedRange);
                }

                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }
            }
        }

        Marshal.ReleaseComObject(sheets);
    }

    private void ScanRangeForDebugCallSites(dynamic usedRange)
    {
        var rows = (int)usedRange.Rows.Count;
        var cols = (int)usedRange.Columns.Count;

        for (var r = 1; r <= rows; r++)
        {
            for (var c = 1; c <= cols; c++)
            {
                dynamic? cell = null;
                try
                {
                    cell = usedRange.Cells[r, c];
                    var formula = cell.Formula2 as string;
                    if (string.IsNullOrEmpty(formula) || !formula!.Contains(CodeEmitter.DebugSuffix))
                    {
                        continue;
                    }

                    var debugNames = LetFormulaReconstructor.GetDebugCallSites(formula);
                    if (debugNames.Count == 0)
                    {
                        continue;
                    }

                    Debug.WriteLine($"Rehydrating debug variants for: {string.Join(", ", debugNames)}");
                    RehydrateCellDebugVariants(formula, debugNames);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Rehydration error at cell [{r},{c}]: {ex.Message}");
                }
                finally
                {
                    if (cell != null)
                    {
                        Marshal.ReleaseComObject(cell);
                    }
                }
            }
        }
    }

    /// <summary>
    ///     Compiles the debug variant for each debug call site found in the formula
    ///     by extracting the DSL source from the matching _src_ bindings.
    /// </summary>
    private void RehydrateCellDebugVariants(string formula, List<string> debugNames)
    {
        if (!LetFormulaParser.TryParse(formula, out var structure) || structure == null)
        {
            return;
        }

        // Build source map from _src_ bindings
        var sourceMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var binding in structure.Bindings)
        {
            var varName = binding.VariableName.Trim();
            if (varName.StartsWith("_src_", StringComparison.Ordinal))
            {
                var targetName = varName["_src_".Length..];
                var value = binding.Value.Trim();
                if (value.StartsWith('"') && value.EndsWith('"') && value.Length >= 2)
                {
                    value = value[1..^1].Replace("\"\"", "\"");
                }

                sourceMap[targetName] = value;
            }
        }

        // Compile both normal and debug variants for each name
        foreach (var name in debugNames)
        {
            if (sourceMap.TryGetValue(name, out var source))
            {
                try
                {
                    var context = new ExpressionContext(name);
                    _pipeline!.Process(source, context);
                    Debug.WriteLine($"Rehydrated debug variant for: {name}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to rehydrate debug variant for {name}: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    ///     COM add-in registered solely to receive Excel shutdown notification.
    /// </summary>
    private class ShutdownMonitor : ExcelComAddIn
    {
        public override void OnBeginShutdown(ref Array custom)
        {
            if (_instance != null)
            {
                _instance._isExcelShutdown = true;
                _instance.Dispose();
            }
        }
    }
}
