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

            // Only process single-cell changes — backtick formulas are entered one cell at a time
            if (target.Cells.CountLarge != 1)
            {
                return;
            }

            var cellText = target.Formula as string;
            if (!BacktickExtractor.IsBacktickFormula(cellText))
            {
                return;
            }

            // Unwrap immediately so cell doesn't stay tall while waiting for processing
            target.WrapText = false;
            var address = target.Address as string;

            // Get worksheet reference for later
            var worksheet = target.Worksheet;

            // Queue processing as a macro - registration must happen in macro context
            ExcelAsyncUtil.QueueAsMacro(() =>
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
            });
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
            var cellText = cell.Formula as string;

            if (!BacktickExtractor.IsBacktickFormula(cellText))
            {
                return;
            }

            Debug.WriteLine($"Found backtick formula: {cellText}");

            var originalFormula = cell.Value2 as string ?? cellText;
            if (originalFormula == null)
            {
                return;
            }

            // Check if this is a LET formula - handle specially for named UDFs
            if (LetFormulaParser.TryParse(originalFormula, out var letStructure) &&
                (letStructure!.Bindings.Any(b => b.HasBacktick) ||
                 letStructure.ResultExpression.Contains('`')))
            {
                ProcessLetFormula(cell, letStructure);
                return;
            }

            // Non-LET formula - use existing backtick processing
            ProcessBacktickFormula(cell, originalFormula);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ProcessCell error: {ex.Message}");
            SetCellError(cell, $"Internal error: {ex.Message}");
        }
    }

    private void ProcessLetFormula(dynamic cell, LetStructure letStructure)
    {
        Debug.WriteLine("Processing LET formula with backtick expressions");

        var processedBindings = new Dictionary<string, ProcessedBinding>();
        var errors = new List<string>();

        foreach (var binding in letStructure.Bindings)
        {
            var variableName = binding.VariableName.Trim();

            if (binding.HasBacktick)
            {
                var dslExpression = LetFormulaParser.ExtractBacktickExpression(binding.Value);
                if (dslExpression == null)
                {
                    errors.Add($"Could not extract backtick expression from {variableName}");
                    continue;
                }

                Debug.WriteLine($"Processing LET binding: {variableName} = `{dslExpression}`");

                var context = new ExpressionContext(variableName);
                var result = _pipeline.Process(dslExpression, context);

                if (result.Success && result.UdfName != null)
                {
                    processedBindings[variableName] = new ProcessedBinding(
                        variableName,
                        dslExpression,
                        result.UdfName,
                        result.Parameters ?? Array.Empty<string>());

                    Debug.WriteLine(
                        $"LET UDF generated: {result.UdfName}({string.Join(", ", result.Parameters ?? Array.Empty<string>())})");
                }
                else
                {
                    errors.Add($"{variableName}: {result.ErrorMessage ?? "Unknown error"}");
                    Debug.WriteLine($"Pipeline error for {variableName}: {result.ErrorMessage}");
                }
            }
        }

        // Check if the result expression also contains a backtick
        ProcessedBinding? processedResult = null;
        if (letStructure.ResultExpression.Contains('`'))
        {
            var dslExpression = LetFormulaParser.ExtractBacktickExpression(letStructure.ResultExpression);
            if (dslExpression != null)
            {
                Debug.WriteLine($"Processing LET result expression: `{dslExpression}`");

                var context = new ExpressionContext("_result");
                var result = _pipeline.Process(dslExpression, context);

                if (result.Success && result.UdfName != null)
                {
                    processedResult = new ProcessedBinding(
                        "_result",
                        dslExpression,
                        result.UdfName,
                        result.Parameters ?? Array.Empty<string>());

                    Debug.WriteLine(
                        $"LET result UDF generated: {result.UdfName}({string.Join(", ", result.Parameters ?? Array.Empty<string>())})");
                }
                else
                {
                    errors.Add($"Result expression: {result.ErrorMessage ?? "Unknown error"}");
                    Debug.WriteLine($"Pipeline error for result expression: {result.ErrorMessage}");
                }
            }
        }

        // Validate: column access on LET variables is not supported
        var letVariableNames = new HashSet<string>(
            letStructure.Bindings
                .Where(b => !b.HasBacktick)
                .Select(b => b.VariableName.Trim()),
            StringComparer.OrdinalIgnoreCase);

        foreach (var (_, processed) in processedBindings)
        {
            foreach (var param in processed.Parameters)
            {
                if (!param.EndsWith("[#All]"))
                {
                    continue;
                }

                var baseName = param[..^"[#All]".Length];
                if (letVariableNames.Contains(baseName))
                {
                    errors.Add(
                        $"Column access (r[\"..\"]) requires a direct table reference. " +
                        $"LET variable '{baseName}' cannot be used — use the table name directly.");
                }
            }
        }

        if (processedResult != null)
        {
            foreach (var param in processedResult.Parameters)
            {
                if (!param.EndsWith("[#All]"))
                {
                    continue;
                }

                var baseName = param[..^"[#All]".Length];
                if (letVariableNames.Contains(baseName))
                {
                    errors.Add(
                        $"Column access (r[\"..\"]) requires a direct table reference. " +
                        $"LET variable '{baseName}' cannot be used — use the table name directly.");
                }
            }
        }

        if (errors.Count > 0)
        {
            SetCellError(cell, string.Join("\n", errors));
            return;
        }

        var newFormula = LetFormulaRewriter.Rewrite(letStructure, processedBindings, processedResult);
        Debug.WriteLine($"Rewriting LET formula to: {newFormula}");

        WriteFormula(cell, newFormula);
    }

    private void ProcessBacktickFormula(dynamic cell, string originalFormula)
    {
        var expressions = BacktickExtractor.Extract(originalFormula);
        if (expressions.Count == 0)
        {
            return;
        }

        var replacements = new Dictionary<string, string>();
        var errors = new List<string>();

        foreach (var expr in expressions)
        {
            Debug.WriteLine($"Processing expression: {expr.Expression}");
            var result = _pipeline.Process(expr.Expression);

            if (result.Success && result.UdfName != null)
            {
                // Build UDF call with all parameters
                var paramStr = result.Parameters != null && result.Parameters.Count > 0
                    ? string.Join(", ", result.Parameters)
                    : "";
                var udfCall = $"{result.UdfName}({paramStr})";
                replacements[expr.Expression] = udfCall;
                Debug.WriteLine($"UDF generated: {udfCall}");
            }
            else
            {
                errors.Add(result.ErrorMessage ?? "Unknown error");
                Debug.WriteLine($"Pipeline error: {result.ErrorMessage}");
            }
        }

        if (errors.Count > 0)
        {
            SetCellError(cell, string.Join("\n", errors));
            return;
        }

        var newFormula = BacktickExtractor.RewriteFormula(originalFormula, replacements);
        Debug.WriteLine($"Rewriting formula to: {newFormula}");

        WriteFormula(cell, newFormula);
    }

    private static void WriteFormula(dynamic cell, string formula)
    {
        try
        {
            cell.Formula2 = formula;
            ClearCellComment(cell);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to write formula: {ex.Message}");
            SetCellError(cell, $"Could not write formula: {ex.Message}");
        }
    }

    private static void SetCellError(dynamic cell, string errorMessage)
    {
        try
        {
            ClearCellComment(cell);
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
