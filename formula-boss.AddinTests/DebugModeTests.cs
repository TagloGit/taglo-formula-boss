using System.Runtime.InteropServices;

using Xunit;
using Xunit.Abstractions;

namespace FormulaBoss.AddinTests;

/// <summary>
///     End-to-end tests for trace debugging: toggle debug mode, assert FB.LastTrace() output,
///     workbook reopen rehydration, and re-edit while debugging.
/// </summary>
[Collection("Excel Addin")]
public class DebugModeTests
{
    /// <summary>
    ///     A scoring-loop expression that iterates over a scores range, tracking per-player
    ///     scores and turn counts with an if/else branch. Produces trace columns:
    ///     currentPlayer, jScore, jTurns, kScore, kTurns, s, branch, kind, depth.
    /// </summary>
    private const string ScoringLoopExpression =
        "{ " +
        "var jScore = 0.0; var jTurns = 0; " +
        "var kScore = 0.0; var kTurns = 0; " +
        "foreach (var s in scores) { " +
        "  var currentPlayer = (jTurns + kTurns) % 2 == 0 ? \"J\" : \"K\"; " +
        "  if (currentPlayer == \"J\") { jScore += (double)s; jTurns++; } " +
        "  else { kScore += (double)s; kTurns++; } " +
        "} " +
        "return jScore - kScore; " +
        "}";

    private readonly ExcelAddinFixture _excel;
    private readonly ITestOutputHelper _output;

    public DebugModeTests(ExcelAddinFixture excel, ITestOutputHelper output)
    {
        _excel = excel;
        _output = output;
    }

    [Fact]
    public void ScoringLoop_TogglesAndTraces()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up scores data
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);
            TestUtilities.SetCellValue(ws, "A4", 40.0);

            // Enter the scoring-loop formula as a LET with backtick
            TestUtilities.EnterBacktickFormula(ws, "B1",
                $"=LET(scores, A1:A4, result, `{ScoringLoopExpression}`, result)");

            // Wait for the normal UDF to compile and return a result
            var result = TestUtilities.WaitForResult(ws, "B1", _output);
            Assert.NotNull(result);

            var formula = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"B1 formula after compile: {formula}");
            _output.WriteLine($"B1 value: {result}");

            // Toggle debug ON by rewriting call sites to _DEBUG
            var cell = ws.Range["B1"];
            try
            {
                ToggleDebugOn(cell);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }

            // Wait for recalc with the debug variant
            Thread.Sleep(3000);

            // Check FB.LastTrace() spills a table
            SetFormula(ws, "D1", "=FB.LastTrace()");
            Thread.Sleep(3000);

            var traceHeader = TestUtilities.GetCellValue(ws, "D1");
            _output.WriteLine($"D1 (trace header[0]): {traceHeader}");

            // The first cell of the trace should be the "kind" header
            Assert.NotNull(traceHeader);
            Assert.Equal("kind", traceHeader);

            // Verify expected columns exist in the header row
            // Headers: kind, depth, branch, then locals in first-seen order, then return
            var headers = ReadTraceHeaderRow(ws, "D1", 15);
            _output.WriteLine($"Trace headers: {string.Join(", ", headers)}");

            Assert.Contains("kind", headers);
            Assert.Contains("depth", headers);
            Assert.Contains("branch", headers);
            Assert.Contains("currentPlayer", headers);
            Assert.Contains("jScore", headers);
            Assert.Contains("jTurns", headers);
            Assert.Contains("kScore", headers);
            Assert.Contains("kTurns", headers);
            Assert.Contains("s", headers);

            // Verify we have data rows (at least entry + 4 iterations + return)
            var secondRowKind = TestUtilities.GetCellValue(ws, "D2") as string;
            _output.WriteLine($"D2 (first data kind): {secondRowKind}");
            Assert.Equal("entry", secondRowKind);
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ToggleOff_RevertsCallSite()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            // Set up data and enter formula
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);

            TestUtilities.EnterBacktickFormula(ws, "B1",
                $"=LET(scores, A1:A2, result, `{ScoringLoopExpression}`, result)");

            var result = TestUtilities.WaitForResult(ws, "B1", _output);
            Assert.NotNull(result);

            // Toggle debug ON
            var cell = ws.Range["B1"];
            try
            {
                ToggleDebugOn(cell);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }

            Thread.Sleep(3000);

            // Verify debug is on
            var debugFormula = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"Debug ON formula: {debugFormula}");
            Assert.Contains("_DEBUG", debugFormula);

            // Check FB.LastTrace() has data
            SetFormula(ws, "D1", "=FB.LastTrace()");
            Thread.Sleep(3000);
            var traceBeforeToggleOff = TestUtilities.GetCellValue(ws, "D1");
            _output.WriteLine($"Trace before toggle-off: {traceBeforeToggleOff}");
            Assert.Equal("kind", traceBeforeToggleOff);

            // Toggle debug OFF
            cell = ws.Range["B1"];
            try
            {
                ToggleDebugOff(cell);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }

            Thread.Sleep(2000);

            // Verify call site reverted (no _DEBUG)
            var normalFormula = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"Debug OFF formula: {normalFormula}");
            Assert.DoesNotContain("_DEBUG", normalFormula);
            Assert.Contains("__FB_", normalFormula);

            // FB.LastTrace() should still show the last captured buffer
            // (buffer is not cleared on toggle-off)
            Thread.Sleep(1000);
            var traceAfterToggleOff = TestUtilities.GetCellValue(ws, "D1");
            _output.WriteLine($"Trace after toggle-off: {traceAfterToggleOff}");
            Assert.Equal("kind", traceAfterToggleOff);
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    [Fact]
    public void ReopenWorkbook_RehydratesDebugVariant()
    {
        // Step 1: Create a workbook with a debug call site, save it
        var tempPath = Path.Combine(Path.GetTempPath(), $"FB_DebugRehydrate_{Guid.NewGuid():N}.xlsx");
        dynamic? newWb = null;

        try
        {
            // Enter formula and compile in the current workbook
            var ws = _excel.AddWorksheet();
            try
            {
                TestUtilities.SetCellValue(ws, "A1", 10.0);
                TestUtilities.SetCellValue(ws, "A2", 20.0);
                TestUtilities.SetCellValue(ws, "A3", 30.0);
                TestUtilities.SetCellValue(ws, "A4", 40.0);

                TestUtilities.EnterBacktickFormula(ws, "B1",
                    $"=LET(scores, A1:A4, result, `{ScoringLoopExpression}`, result)");

                var result = TestUtilities.WaitForResult(ws, "B1", _output);
                Assert.NotNull(result);

                // Toggle debug ON
                var cell = ws.Range["B1"];
                try
                {
                    ToggleDebugOn(cell);
                }
                finally
                {
                    Marshal.ReleaseComObject(cell);
                }

                Thread.Sleep(3000);

                // Verify _DEBUG call site is present
                var debugFormula = TestUtilities.GetCellFormula(ws, "B1");
                _output.WriteLine($"Debug formula before save: {debugFormula}");
                Assert.Contains("_DEBUG", debugFormula);

                // Copy data and formula to a new workbook and save it
                newWb = _excel.Application.Workbooks.Add();
                var newWs = newWb.Worksheets[1];
                try
                {
                    // Copy data cells
                    TestUtilities.SetCellValue(newWs, "A1", 10.0);
                    TestUtilities.SetCellValue(newWs, "A2", 20.0);
                    TestUtilities.SetCellValue(newWs, "A3", 30.0);
                    TestUtilities.SetCellValue(newWs, "A4", 40.0);

                    // Copy the debug formula directly
                    var newCell = newWs.Range["B1"];
                    try
                    {
                        newCell.Formula2 = debugFormula;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(newCell);
                    }

                    // Save as xlsx
                    // xlOpenXMLWorkbook = 51
                    newWb.SaveAs(tempPath, 51);
                }
                finally
                {
                    Marshal.ReleaseComObject(newWs);
                }

                // Close the saved workbook
                newWb.Close(false);
                Marshal.ReleaseComObject(newWb);
                newWb = null;
            }
            finally
            {
                TestUtilities.CleanupWorksheet(ws);
            }

            // Step 2: Reopen the saved workbook — rehydration should compile the debug variant
            newWb = _excel.Application.Workbooks.Open(tempPath);
            Thread.Sleep(5000); // Wait for WorkbookOpen event + rehydration + recalc

            var reopenWs = newWb.Worksheets[1];
            try
            {
                var reopenFormula = TestUtilities.GetCellFormula(reopenWs, "B1");
                var reopenValue = TestUtilities.GetCellValue(reopenWs, "B1");
                _output.WriteLine($"After reopen - formula: {reopenFormula}");
                _output.WriteLine($"After reopen - value: {reopenValue} (type: {reopenValue?.GetType()?.Name})");

                // Formula should still have _DEBUG call site
                Assert.Contains("_DEBUG", reopenFormula);

                // Value should NOT be #NAME? — the debug variant should be compiled
                Assert.NotNull(reopenValue);
                var isNameError = reopenValue is int errorCode && errorCode == -2146826259;
                var isStringError = reopenValue is string s && s.Contains("#NAME");
                Assert.False(isNameError || isStringError,
                    $"Expected valid result but got: {reopenValue}");

                // Verify FB.LastTrace() works after rehydration
                SetFormula(reopenWs, "D1", "=FB.LastTrace()");
                Thread.Sleep(3000);

                var traceValue = TestUtilities.GetCellValue(reopenWs, "D1");
                _output.WriteLine($"Trace after reopen: {traceValue}");
                Assert.Equal("kind", traceValue);
            }
            finally
            {
                Marshal.ReleaseComObject(reopenWs);
            }
        }
        finally
        {
            if (newWb != null)
            {
                try
                {
                    newWb.Close(false);
                    Marshal.ReleaseComObject(newWb);
                }
                catch
                {
                    // Ignore
                }
            }

            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
                // Ignore
            }
        }
    }

    [Fact]
    public void EditSourceWhileDebugging_RecompilesBoth()
    {
        var ws = _excel.AddWorksheet();
        try
        {
            TestUtilities.SetCellValue(ws, "A1", 10.0);
            TestUtilities.SetCellValue(ws, "A2", 20.0);
            TestUtilities.SetCellValue(ws, "A3", 30.0);
            TestUtilities.SetCellValue(ws, "A4", 40.0);

            // Enter initial formula
            TestUtilities.EnterBacktickFormula(ws, "B1",
                $"=LET(scores, A1:A4, result, `{ScoringLoopExpression}`, result)");

            var result1 = TestUtilities.WaitForResult(ws, "B1", _output);
            Assert.NotNull(result1);
            _output.WriteLine($"Initial result: {result1}");

            // Toggle debug ON
            var cell = ws.Range["B1"];
            try
            {
                ToggleDebugOn(cell);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }

            Thread.Sleep(3000);

            var debugFormula1 = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"Debug formula 1: {debugFormula1}");
            Assert.Contains("_DEBUG", debugFormula1);

            // Edit source: enter a different expression (simulates re-editing)
            // Use a simpler expression that still loops
            const string newExpression =
                "{ var total = 0.0; foreach (var s in scores) { total += (double)s; } return total; }";
            TestUtilities.EnterBacktickFormula(ws, "B1",
                $"=LET(scores, A1:A4, result, `{newExpression}`, result)");

            // Wait for recompile of both normal and debug variants
            var result2 = TestUtilities.WaitForResult(ws, "B1", _output);
            Assert.NotNull(result2);
            _output.WriteLine($"After edit result: {result2}");

            // The re-entered formula goes back through the normal pipeline (backtick text),
            // which compiles both normal + debug. The call site won't have _DEBUG anymore
            // (user re-entered the backtick formula, so interceptor rewrites to normal).
            // But toggling debug back on should be instant since both variants are compiled.
            var formulaAfterEdit = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"Formula after edit: {formulaAfterEdit}");

            // Toggle debug ON again — should be instant since debug variant was compiled
            cell = ws.Range["B1"];
            try
            {
                ToggleDebugOn(cell);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }

            Thread.Sleep(3000);

            // Verify it's in debug mode and produces a result (not #NAME?)
            var debugFormula2 = TestUtilities.GetCellFormula(ws, "B1");
            _output.WriteLine($"Debug formula 2: {debugFormula2}");
            Assert.Contains("_DEBUG", debugFormula2);

            var debugValue = TestUtilities.GetCellValue(ws, "B1");
            _output.WriteLine($"Debug value after re-toggle: {debugValue} (type: {debugValue?.GetType()?.Name})");
            Assert.NotNull(debugValue);

            // Should compute total: 10 + 20 + 30 + 40 = 100
            Assert.Equal(100.0, Convert.ToDouble(debugValue));

            // Verify trace works with new expression
            SetFormula(ws, "D1", "=FB.LastTrace()");
            Thread.Sleep(3000);
            var traceValue = TestUtilities.GetCellValue(ws, "D1");
            _output.WriteLine($"Trace after re-edit: {traceValue}");
            Assert.Equal("kind", traceValue);
        }
        finally
        {
            TestUtilities.CleanupWorksheet(ws);
        }
    }

    /// <summary>
    ///     Toggles debug ON for the given cell by rewriting call sites from
    ///     __FB_NAME( to __FB_NAME_DEBUG(.
    /// </summary>
    private void ToggleDebugOn(dynamic cell)
    {
        var formula = cell.Formula2 as string ?? "";
        _output.WriteLine($"ToggleDebugOn input: {formula}");

        // Find normal call sites and rewrite to debug
        var prefix = "__FB_";
        var names = new List<string>();
        var searchFrom = 0;
        while (true)
        {
            var idx = formula.IndexOf(prefix, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                break;
            }

            var nameStart = idx + prefix.Length;
            var parenIdx = formula.IndexOf('(', nameStart);
            if (parenIdx < 0)
            {
                break;
            }

            var name = formula[nameStart..parenIdx];
            if (!name.EndsWith("_DEBUG", StringComparison.OrdinalIgnoreCase))
            {
                names.Add(name);
            }

            searchFrom = parenIdx + 1;
        }

        if (names.Count == 0)
        {
            _output.WriteLine("ToggleDebugOn: no normal call sites found");
            return;
        }

        // Rewrite call sites (outside string literals) to _DEBUG
        var result = formula;
        foreach (var name in names)
        {
            result = result.Replace($"__FB_{name}(", $"__FB_{name}_DEBUG(");
        }

        _output.WriteLine($"ToggleDebugOn output: {result}");
        cell.Formula2 = result;
    }

    /// <summary>
    ///     Toggles debug OFF for the given cell by rewriting call sites from
    ///     __FB_NAME_DEBUG( to __FB_NAME(.
    /// </summary>
    private void ToggleDebugOff(dynamic cell)
    {
        var formula = cell.Formula2 as string ?? "";
        _output.WriteLine($"ToggleDebugOff input: {formula}");

        // Find _DEBUG call sites and rewrite to normal
        var result = formula;
        var prefix = "__FB_";
        var searchFrom = 0;
        while (true)
        {
            var idx = result.IndexOf(prefix, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                break;
            }

            var nameStart = idx + prefix.Length;
            var parenIdx = result.IndexOf('(', nameStart);
            if (parenIdx < 0)
            {
                break;
            }

            var name = result[nameStart..parenIdx];
            if (name.EndsWith("_DEBUG", StringComparison.OrdinalIgnoreCase))
            {
                var baseName = name[..^"_DEBUG".Length];
                result = result.Replace($"__FB_{name}(", $"__FB_{baseName}(");
                // Don't advance searchFrom — replacement shortened the string
                continue;
            }

            searchFrom = parenIdx + 1;
        }

        _output.WriteLine($"ToggleDebugOff output: {result}");
        cell.Formula2 = result;
    }

    /// <summary>
    ///     Reads the header row of a trace spill starting at the given cell.
    /// </summary>
    private List<string> ReadTraceHeaderRow(dynamic ws, string startCell, int maxCols)
    {
        var headers = new List<string>();
        // Parse start cell to get column letter and row number
        var col = 0;
        var row = 0;
        foreach (var ch in startCell)
        {
            if (char.IsLetter(ch))
            {
                col = (col * 26) + (char.ToUpper(ch) - 'A') + 1;
            }
            else
            {
                row = (row * 10) + (ch - '0');
            }
        }

        for (var c = col; c < col + maxCols; c++)
        {
            var cell = ws.Cells[row, c];
            try
            {
                var value = cell.Value;
                if (value == null)
                {
                    break;
                }

                headers.Add(value.ToString()!);
            }
            finally
            {
                Marshal.ReleaseComObject(cell);
            }
        }

        return headers;
    }

    /// <summary>
    ///     Sets a cell's Formula2 directly (not as text — used for UDF formulas like FB.LastTrace).
    /// </summary>
    private static void SetFormula(dynamic ws, string cellAddress, string formula)
    {
        var cell = ws.Range[cellAddress];
        try
        {
            cell.Formula2 = formula;
        }
        finally
        {
            Marshal.ReleaseComObject(cell);
        }
    }
}
