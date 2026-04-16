using System.Diagnostics;
using System.Runtime.InteropServices;

using ExcelDna.Integration;

using FormulaBoss.Transpilation;
using FormulaBoss.UI;

using Taglo.Excel.Common;

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

            try
            {
                Marshal.ReleaseComObject(_app);
            }
            catch
            {
                // Ignore — may already be released
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
                    Logger.Error($"ProcessCell({address})", ex);
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Error("OnSheetChange", ex);
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

            // Capture workbook metadata for metadata-aware header detection
            var metadata = _app != null ? WorkbookMetadata.CaptureFromExcel(_app) : null;

            // Check if this is a LET formula - handle specially for named UDFs
            if (LetFormulaParser.TryParse(originalFormula, out var letStructure) &&
                (letStructure!.Bindings.Any(b => b.HasBacktick) ||
                 letStructure.ResultExpression.Contains('`')))
            {
                ProcessLetFormula(cell, letStructure, metadata);
                return;
            }

            // Non-LET formula - use existing backtick processing
            ProcessBacktickFormula(cell, originalFormula, metadata);
        }
        catch (Exception ex)
        {
            Logger.Error("ProcessCell", ex);
            SetCellError(cell, $"Internal error: {ex.Message}");
        }
    }

    private void ProcessLetFormula(dynamic cell, LetStructure letStructure, WorkbookMetadata? metadata)
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

                var context = new ExpressionContext(variableName, metadata);
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

        // Check if the result expression also contains backtick(s)
        var processedResults = new List<ProcessedBinding>();
        string? rewrittenResultExpression = null;
        if (letStructure.ResultExpression.Contains('`'))
        {
            var backtickExprs = BacktickExtractor.Extract(letStructure.ResultExpression);
            for (var i = 0; i < backtickExprs.Count; i++)
            {
                var dslExpression = backtickExprs[i].Expression;
                var varName = backtickExprs.Count == 1 ? "_result" : $"_result_{i + 1}";

                Debug.WriteLine($"Processing LET result expression: `{dslExpression}`");

                var context = new ExpressionContext(varName, metadata);
                var result = _pipeline.Process(dslExpression, context);

                if (result.Success && result.UdfName != null)
                {
                    processedResults.Add(new ProcessedBinding(
                        varName,
                        dslExpression,
                        result.UdfName,
                        result.Parameters ?? Array.Empty<string>()));

                    Debug.WriteLine(
                        $"LET result UDF generated: {result.UdfName}({string.Join(", ", result.Parameters ?? Array.Empty<string>())})");
                }
                else
                {
                    errors.Add($"Result expression: {result.ErrorMessage ?? "Unknown error"}");
                    Debug.WriteLine($"Pipeline error for result expression: {result.ErrorMessage}");
                }
            }

            // Build the rewritten result expression by replacing backticks with variable names
            if (processedResults.Count > 0 && processedResults.Count == backtickExprs.Count)
            {
                var replacements = new Dictionary<string, string>();
                for (var i = 0; i < backtickExprs.Count; i++)
                {
                    replacements[backtickExprs[i].Expression] = processedResults[i].VariableName;
                }

                rewrittenResultExpression =
                    BacktickExtractor.RewriteFormula(letStructure.ResultExpression, replacements);
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

        foreach (var processedResult in processedResults)
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

        var settings = EditorSettings.Load();
        var newFormula = LetFormulaRewriter.Rewrite(letStructure, processedBindings, processedResults,
            rewrittenResultExpression, settings.IndentSize, settings.NestedLetDepth, settings.MaxLineLength);
        Debug.WriteLine($"Rewriting LET formula to: {newFormula}");

        WriteFormula(cell, newFormula);
    }

    private void ProcessBacktickFormula(dynamic cell, string originalFormula, WorkbookMetadata? metadata)
    {
        var expressions = BacktickExtractor.Extract(originalFormula);
        if (expressions.Count == 0)
        {
            return;
        }

        var replacements = new Dictionary<string, string>();
        // Track each processed expression for _src_ binding generation (preserves order)
        var processedExpressions = new List<(string DslExpression, string UdfName, string UdfCall)>();
        var errors = new List<string>();

        foreach (var expr in expressions)
        {
            Debug.WriteLine($"Processing expression: {expr.Expression}");
            var context = metadata != null ? new ExpressionContext(null, metadata) : null;
            var result = _pipeline.Process(expr.Expression, context);

            if (result.Success && result.UdfName != null)
            {
                // Build UDF call with all parameters
                var paramStr = result.Parameters != null && result.Parameters.Count > 0
                    ? string.Join(", ", result.Parameters)
                    : "";
                var udfCall = $"{result.UdfName}({paramStr})";
                replacements[expr.Expression] = udfCall;
                processedExpressions.Add((expr.Expression, result.UdfName, udfCall));
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

        // Rewrite backtick expressions to UDF calls
        var rewrittenFormula = BacktickExtractor.RewriteFormula(originalFormula, replacements);

        // Wrap in LET with _src_ bindings so the formula is rehydration-capable and editor-reopenable
        var newFormula = WrapInLetWithSourceBindings(rewrittenFormula, processedExpressions);
        Debug.WriteLine($"Rewriting formula to: {newFormula}");

        WriteFormula(cell, newFormula);
    }

    /// <summary>
    ///     Wraps a rewritten formula in a LET with _src_ bindings for each UDF,
    ///     making plain backtick formulas compatible with rehydration and editor reopen.
    /// </summary>
    private static string WrapInLetWithSourceBindings(
        string rewrittenFormula,
        List<(string DslExpression, string UdfName, string UdfCall)> processedExpressions)
    {
        // Strip leading = for use as the LET result expression
        var resultExpression = rewrittenFormula.StartsWith('=')
            ? rewrittenFormula[1..]
            : rewrittenFormula;

        var sb = new System.Text.StringBuilder();
        sb.Append("=LET(");

        foreach (var (dslExpression, udfName, _) in processedExpressions)
        {
            // Strip __FB_ prefix for the _src_ key so rehydration can match it
            // (GetDebugCallSites captures the name between __FB_ and _DEBUG)
            var srcKey = udfName.StartsWith(CodeEmitter.UdfPrefix, StringComparison.OrdinalIgnoreCase)
                ? udfName[CodeEmitter.UdfPrefix.Length..]
                : udfName;
            sb.Append("_src_").Append(srcKey).Append(", ");
            sb.Append('"').Append(LetFormulaRewriter.EscapeForExcelString(dslExpression)).Append("\", ");
        }

        sb.Append(resultExpression);
        sb.Append(')');

        return sb.ToString();
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
