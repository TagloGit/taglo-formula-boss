using System.Diagnostics;

using ExcelDna.Integration;

namespace FormulaBoss.Interception;

/// <summary>
///     Intercepts worksheet changes to detect and process backtick formulas.
/// </summary>
public class FormulaInterceptor : IDisposable
{
    private readonly FormulaPipeline _pipeline;
    private dynamic? _app;
    private bool _disposed;
    private bool _isProcessing;

    public FormulaInterceptor(FormulaPipeline pipeline)
    {
        _pipeline = pipeline;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Stop();
            _disposed = true;
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Starts listening for worksheet changes.
    /// </summary>
    public void Start()
    {
        _app = ExcelDnaUtil.Application;
        if (_app != null)
        {
            _app.SheetChange += new SheetChangeHandler(OnSheetChange);
            Debug.WriteLine("FormulaInterceptor: SheetChange event hooked");
        }
        else
        {
            Debug.WriteLine("FormulaInterceptor: Could not get Excel Application");
        }
    }

    /// <summary>
    ///     Stops listening for worksheet changes.
    /// </summary>
    public void Stop()
    {
        if (_app != null)
        {
            try
            {
                _app.SheetChange -= new SheetChangeHandler(OnSheetChange);
            }
            catch
            {
                // Ignore errors during cleanup
            }

            _app = null;
        }
    }

    private void OnSheetChange(object sheet, dynamic target)
    {
        // Prevent re-entrancy when we modify the cell
        if (_isProcessing)
        {
            return;
        }

        try
        {
            _isProcessing = true;

            // Collect cells that need processing
            var cellsToProcess = new List<string>();
            foreach (var cell in target.Cells)
            {
                var cellText = cell.Formula as string;
                if (BacktickExtractor.IsBacktickFormula(cellText))
                {
                    // Store the cell address to process later
                    var address = cell.Address as string;
                    if (address != null)
                    {
                        cellsToProcess.Add(address);
                    }
                }
            }

            if (cellsToProcess.Count > 0)
            {
                // Get worksheet reference for later
                dynamic worksheet = target.Worksheet;

                // Queue processing as a macro - registration must happen in macro context
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    foreach (var address in cellsToProcess)
                    {
                        try
                        {
                            var cell = worksheet.Range[address];
                            ProcessCell(cell);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"ProcessCell error for {address}: {ex.Message}");
                        }
                    }
                });
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"OnSheetChange error: {ex.Message}");
        }
        finally
        {
            _isProcessing = false;
        }
    }

    private void ProcessCell(dynamic cell)
    {
        try
        {
            // Get the cell's text (what the user typed, shown as text due to quote prefix)
            var cellText = cell.Formula as string;

            // Check if it's a backtick formula
            if (!BacktickExtractor.IsBacktickFormula(cellText))
            {
                return;
            }

            Debug.WriteLine($"Found backtick formula: {cellText}");

            // The cell shows as text because of the leading apostrophe
            // Get the actual value which includes the apostrophe
            var originalFormula = cell.Value2 as string ?? cellText;
            if (originalFormula == null)
            {
                return;
            }

            // Extract backtick expressions
            var expressions = BacktickExtractor.Extract(originalFormula);
            if (expressions.Count == 0)
            {
                return;
            }

            // Process each expression and build replacements
            var replacements = new Dictionary<string, string>();
            var errors = new List<string>();

            foreach (var expr in expressions)
            {
                Debug.WriteLine($"Processing expression: {expr.Expression}");
                var result = _pipeline.Process(expr.Expression);

                if (result.Success && result.UdfName != null)
                {
                    // Build the UDF call with the input parameter
                    var udfCall = $"{result.UdfName}({result.InputParameter})";
                    replacements[expr.Expression] = udfCall;
                    Debug.WriteLine($"UDF generated: {udfCall}");
                }
                else
                {
                    errors.Add(result.ErrorMessage ?? "Unknown error");
                    Debug.WriteLine($"Pipeline error: {result.ErrorMessage}");
                }
            }

            // If there were errors, add a comment and leave the cell as-is
            if (errors.Count > 0)
            {
                SetCellError(cell, string.Join("\n", errors));
                return;
            }

            // Rewrite the formula
            var newFormula = BacktickExtractor.RewriteFormula(originalFormula, replacements);
            Debug.WriteLine($"Rewriting formula to: {newFormula}");

            // Set the cell formula (this will trigger another SheetChange, but _isProcessing guards us)
            cell.Formula = newFormula;

            // Clear any previous error comment
            ClearCellComment(cell);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ProcessCell error: {ex.Message}");
            SetCellError(cell, $"Internal error: {ex.Message}");
        }
    }

    private static void SetCellError(dynamic cell, string errorMessage)
    {
        try
        {
            // Clear existing comment
            ClearCellComment(cell);

            // Add error comment
            cell.AddComment($"Formula Boss Error:\n{errorMessage}");
        }
        catch
        {
            // Ignore errors when setting the comment
        }
    }

    private static void ClearCellComment(dynamic cell)
    {
        try
        {
            cell.Comment?.Delete();
        }
        catch
        {
            // Ignore errors when clearing comment
        }
    }

    // Delegate type for SheetChange event
    private delegate void SheetChangeHandler(object sheet, dynamic target);
}
