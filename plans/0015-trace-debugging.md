# 0015 — Trace Debugging — Implementation Plan

## Overview

Add a **debug mode** to the FB transpile pipeline that emits an instrumented variant of each compiled UDF (`__FB_<NAME>_DEBUG`) alongside the normal one. The debug variant records a per-iteration / per-return snapshot of every in-scope local into a per-cell buffer. A companion UDF `=FB.LastTrace()` spills the most recent buffer as a 2D array.

The toggle is per-cell in the floating editor. Toggling rewrites call sites in the LET formula (`__FB_FOO(...)` ↔ `__FB_FOO_DEBUG(...)`) without touching the `_src_foo` source string literal. Because the cell's call site is the mode indicator, debug mode survives file reopen with no hidden state.

The implementation splits cleanly into five PR-sized tasks that stack on each other: (1) the `Tracer` runtime bridge, (2) the `DebugInstrumentationRewriter` that emits snapshot calls, (3) the compile/register path for `_DEBUG` variants plus `FB.LastTrace`, (4) the editor toggle UI and call-site rewriter, and (5) workbook-open rehydration + integration tests.

## Files to Touch

### New files

- `formula-boss.Runtime/Tracer.cs` — buffer + delegate bridge class. Public static API: `Begin(name, callerAddr)`, `Set(name, value)`, `Snapshot(kind, depth, branch)`, `Return(value)`, `TruncateWarn()`. Static `LastBuffer` keyed by caller address. Delegate fields follow the `RuntimeBridge` pattern (only `object`/primitives in signatures).
- `formula-boss/Transpilation/DebugInstrumentationRewriter.cs` — `CSharpSyntaxRewriter` that wraps the user's statement block: inserts `Tracer.Begin`, `Tracer.Set` after each `LocalDeclarationStatement` and `AssignmentExpressionStatement`, `Tracer.Snapshot("iter", depth, branchLabel)` at the end of each loop body, `Tracer.Snapshot("return", ..)` + return-value capture around each `ReturnStatement`, and branch labelling on `IfStatement`/`ElseClause` arms. Tracks loop depth via a visit counter. Extracts branch labels from the first statement of each arm (truncated to 20 chars).
- `formula-boss/LastTraceUdf.cs` — hosts the `FB.LastTrace()` `[ExcelFunction]` static. Returns `Tracer.LastBuffer.ToObjectArray()` or `"#N/A — no trace captured"` if empty.
- `formula-boss.Tests/DebugInstrumentationRewriterTests.cs` — unit tests asserting emitted source shape for representative inputs (simple foreach, nested foreach, if/else, return inside loop, early return).
- `formula-boss.Tests/TracerTests.cs` — unit tests for buffer behaviour: row cap at 1000 + truncation-warning row, column union across snapshots, cell keying, clear-on-Begin.
- `formula-boss.AddinTests/DebugModeTests.cs` — end-to-end: enter a formula, toggle debug, assert `FB.LastTrace()` spills the expected table; toggle off, assert call site reverts.

### Modified files

- `formula-boss/Transpilation/TranspileResult.cs` — add `DebugSourceCode` field (or a parallel `TranspileResult DebugVariant` property). The normal path is unchanged; debug mode is an explicit second compile pass.
- `formula-boss/Transpilation/CodeEmitter.cs` — add `EmitDebug(...)` path that runs `DebugInstrumentationRewriter` over the user block before emission, and generates the `__FB_<NAME>_DEBUG` method name (new constant `DebugSuffix = "_DEBUG"`).
- `formula-boss/Compilation/DynamicCompiler.cs` — add `CompileAndRegisterDebug(...)` that compiles the debug variant and registers the `_DEBUG` delegate. Called from the pipeline when debug mode is requested for a cell.
- `formula-boss/Interception/LetFormulaReconstructor.cs` — add helpers: `RewriteCallSitesToDebug(formula, names)` and `RewriteCallSitesToNormal(formula, names)` that flip `__FB_<NAME>(` ↔ `__FB_<NAME>_DEBUG(` without touching `_src_<name>` strings. Add `GetDebugCallSites(formula)` returning the names of any `_DEBUG` call sites present (for open-workbook rehydration).
- `formula-boss/Commands/ShowFloatingEditorCommand.cs` — no direct changes; the toggle is wired through the editor view, which calls into a new `DebugToggleService`.
- `formula-boss/UI/Editor/*` (exact file per the editor's current structure) — add a **Debug** toolbar button + `Ctrl+D` binding on the floating editor. The button's `IsChecked` state reflects whether the current cell's formula already has `_DEBUG` call sites.
- `formula-boss/UI/RibbonController.cs` — no change needed (toggle lives in the floating editor, not the ribbon).
- `formula-boss/AddIn.cs` — in `AutoOpen()`, initialise the `Tracer` delegate bridge (same pattern as `RuntimeHelpers` / `RuntimeBridge`). On workbook open, scan active sheets for cells whose formulas contain `__FB_*_DEBUG(` and ensure the debug variant is compiled. (Keep this scan minimal — only cells discovered via a cheap scan of the used range; defer heavy discovery if perf becomes an issue.)
- `formula-boss.dna` — add `Pack="false"` / ensure `formula-boss.Runtime.dll` is shipped so the `Tracer` bridge loads correctly (already the case per CLAUDE.md).

## Order of Operations

1. **Task 1 — `Tracer` runtime bridge.** Add `Tracer.cs` to `formula-boss.Runtime`. Implement the buffer (dictionary of column name → list of values, plus `kind`, `depth`, `branch`, `return` columns), row cap + truncation row, cell-address keying, and `ToObjectArray()` that produces `object[,]` with header row. Wire `Tracer` delegate-bridge initialisation from `AddIn.AutoOpen`. Unit tests in `TracerTests.cs`. **Rationale:** this is the foundation — everything else calls into it.

2. **Task 2 — `DebugInstrumentationRewriter`.** Add the rewriter and `EmitDebug` path in `CodeEmitter`. No compile/register wiring yet — tests assert the emitted *source string* shape. Cover: simple statements, `foreach`, nested `foreach` (depth tracking), `if/else` (branch labels), `return` statements. **Rationale:** pure syntactic transform, easy to test in isolation, biggest correctness risk in the whole feature.

3. **Task 3 — Compile/register `_DEBUG` variants + `FB.LastTrace()` UDF.** Add `DynamicCompiler.CompileAndRegisterDebug`. Add `LastTraceUdf.cs` and register it at add-in startup (same path as any other static `[ExcelFunction]`). At this point you can manually rewrite a cell's formula to `__FB_FOO_DEBUG(...)` in the Excel formula bar and see `=FB.LastTrace()` spill. **Rationale:** end-to-end compile path working before any UI is built — de-risks tasks 4 and 5.

4. **Task 4 — Editor toggle UI + call-site rewriter.** Add **Debug** button + `Ctrl+D` in the floating editor. Wire to a new `DebugToggleService` that: reads the active cell's formula, uses `LetFormulaReconstructor` helpers to flip call sites, triggers compile of the `_DEBUG` variant via task 3's path, and writes the updated formula back. Button state reflects current mode by scanning for `_DEBUG` call sites. **Rationale:** this is the user-visible surface and depends on everything before it.

5. **Task 5 — Workbook-open rehydration + AddinTests.** On `WorkbookOpen`, scan used ranges for `__FB_*_DEBUG(` and compile the debug variant for each found name (source from the cell's `_src_<name>` literal). Add end-to-end AddinTests using the scoring-loop example from the spec: enter formula, toggle debug, assert `FB.LastTrace()` spill contents match expected rows/cols; edit source while in debug mode, assert both variants recompile; toggle off, assert call site reverts. **Rationale:** closes the loop on "survives file reopen" and provides the acceptance-criteria proof.

## Testing Approach

- **Unit (`formula-boss.Tests`):**
  - `DebugInstrumentationRewriterTests` — assert on emitted source for each construct. Keep assertions loose (regex or "contains `Tracer.Snapshot(\"iter\"`") rather than exact string match, so trivial formatting changes don't break tests.
  - `TracerTests` — buffer mechanics, row cap, truncation row, column union, cell keying, clear-on-Begin.
  - `LetFormulaReconstructorTests` — add cases for `RewriteCallSitesToDebug` / `RewriteCallSitesToNormal` round-tripping, and ensure `_src_` literals are not touched.

- **Integration (`formula-boss.AddinTests`):**
  - `DebugModeTests.ScoringLoop_TogglesAndTraces` — uses the spec's scoring-loop example. Enters the formula, invokes the toggle service (or simulates `Ctrl+D`), asserts `FB.LastTrace()` spills a table with the expected columns and at least the expected first few rows.
  - `DebugModeTests.ToggleOff_RevertsCallSite` — assert call site reverts and `FB.LastTrace()` still shows the last captured buffer (buffer is not cleared on toggle-off).
  - `DebugModeTests.ReopenWorkbook_RehydratesDebugVariant` — write a workbook with a `_DEBUG` call site, close+reopen, assert no `#NAME?` and `FB.LastTrace()` works after recalc.
  - `DebugModeTests.EditSourceWhileDebugging_RecompilesBoth` — edit the source, assert both normal and debug variants are compiled (toggle off is instant).

- **Perf sanity check (manual, noted in task 3 acceptance):** time a compile of a medium expression with and without the debug pass; a regression >2× indicates the rewriter needs attention. No automated perf test.
