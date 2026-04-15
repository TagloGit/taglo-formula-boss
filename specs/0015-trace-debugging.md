# 0015 — Trace Debugging ("LINQPad-style dumps")

## Problem

Formula Boss lets users write inline C# expressions that transpile to UDFs. When a result is wrong, the user has no visibility into intermediate state: local variables, loop iterations, which branch of an `if/else` fired. Real debugger attach (VS on the Excel process) is a power-user workflow that requires tooling outside Excel and breaks the "stay in the sheet" ergonomics.

A typical bug involves iterating over an array while maintaining a handful of counters and branching on conditions — e.g. a scoring loop that tracks per-player scores and turn counts. When the answer is wrong, the user currently has to reason about the code statically, add scratch cells, or rewrite the expression to return intermediate state. All of these lose the loop structure.

## Proposed Solution

A **trace mode** for compiled expressions. The transpiler emits an instrumented variant of the UDF that records a snapshot of every in-scope local at each loop iteration and at each `return`, into a per-cell buffer. The user spills the buffer as a 2D array via a companion UDF `=FB.LastTrace()`.

Debug mode is toggled per cell via a button / hotkey in the floating editor. Toggling rewrites the cell's call site from `__FB_FOO(...)` to `__FB_FOO_DEBUG(...)`, leaving the `_src_foo` source string literal (which the LET formula already carries) untouched. No hidden editor state — the cell's current call site *is* the mode indicator, so debug mode survives file reopen.

### End-to-end workflow

1. User opens the floating editor on a cell whose answer looks wrong.
2. User presses **Debug** (toolbar button or `Ctrl+D`).
3. FB parses the cell formula, finds `_src_<name>` literals, transpiles each in debug mode, compiles and registers `__FB_<NAME>_DEBUG` delegates, and rewrites the cell's call sites to the `_DEBUG` variants.
4. Excel recalcs. The cell still shows its real answer; the tracer has populated a buffer keyed by the calling cell.
5. User types `=FB.LastTrace()` in an empty cell; it spills a table — one row per loop iteration (plus entry and return rows), one column per local variable plus a `branch` column.
6. User identifies the bug by scanning the trace, edits the source, recompiles. While the cell is in debug mode, recompiles keep producing the `_DEBUG` variant.
7. User presses **Debug** again to toggle off. FB rewrites the call site back to `__FB_<NAME>` and drops the debug delegate.

### What the transpiler emits in debug mode

For each compiled expression:
- At entry: `Tracer.Begin(name)` and an initial snapshot tagged `"entry"` capturing every declared local.
- After each local declaration / assignment: record the new value against the local's name.
- At the top of each loop body: bind the loop variable as a tracked local.
- At the end of each loop body: `Tracer.Snapshot("iter")` — copies the current variable map as one row.
- At each `return`: `Tracer.Snapshot("return")` before the return executes, capturing the final state and the returned value.
- Around each `if/else` arm: tag the arm with a short label (the first statement, truncated) so a `branch` column can be populated.

`Tracer` is a bridge class in `FormulaBoss.Runtime` following the existing delegate-bridge pattern (no host-loaded types in its public surface).

### Buffer lifecycle

- One buffer per calling cell (keyed by `xlfCaller` address).
- Cleared at the start of each invocation of a `_DEBUG` UDF.
- `FB.LastTrace()` (no args) returns the most recently populated buffer as `object[,]`.
- Columns are the union of locals ever snapshotted in that run; missing values render as blank.
- Nested loops flatten into a single table with a `depth` column indicating the loop nesting level (0 = outermost).
- `return` rows include the returned value in a dedicated `return` column.
- Hard cap of 1000 snapshot rows per run. On overflow, a final truncation-warning row is appended and further snapshots are dropped.
- Branch labels are derived from the first statement of the arm, truncated to ~20 chars.
- When the user edits a cell that is currently in debug mode, FB recompiles both the normal and `_DEBUG` variants so toggling off is instant.

## User Stories

- As a Formula Boss user, when my expression returns the wrong answer, I want to see the value of every local at every loop iteration so I can find where the state diverges from what I expected.
- As a Formula Boss user, I want to toggle debug mode on one cell without editing the source, so I can investigate without risking changes to a working formula.
- As a Formula Boss user, I want debug mode to survive a file reopen so I can come back to a debugging session without re-setting it up.
- As a Formula Boss user, I want the trace output to appear as a normal spilled array in the sheet so I can filter, sort, and inspect it with familiar Excel tools.
- As a Formula Boss user, I want zero perf cost in non-debug mode, so I can leave production formulas untouched.

## Acceptance Criteria

- [ ] A **Debug** toggle (button + `Ctrl+D`) exists in the floating editor, enabled when the current cell has one or more `_src_<name>` literals.
- [ ] Toggling debug **on** rewrites each `__FB_<NAME>(...)` call site in the cell formula to `__FB_<NAME>_DEBUG(...)`, compiles and registers the debug delegate, and leaves `_src_<name>` strings untouched.
- [ ] Toggling debug **off** rewrites `__FB_<NAME>_DEBUG(...)` back to `__FB_<NAME>(...)`.
- [ ] On workbook open, any cell whose call site references a `_DEBUG` variant causes FB to compile the debug delegate for the referenced `_src_` source.
- [ ] The debug-mode transpiler emits a `Tracer.Begin / Set / Snapshot / Return` scaffold that captures every local and loop variable at: entry, end of each loop body iteration, and every `return`.
- [ ] The `branch` column identifies which `if/else` arm executed at each snapshot row.
- [ ] `=FB.LastTrace()` spills a 2D array: one header row, one row per snapshot, one column per distinct local + `iter`, `kind` (entry/iter/return), and `branch` columns.
- [ ] For the scoring-loop example in the spec, the trace correctly shows the per-iteration values of `currentPlayer`, `jScore`, `jTurns`, `kScore`, `kTurns`, `s`, and the branch taken.
- [ ] Non-debug UDFs have no tracer code and no measurable perf regression vs. before this feature.
- [ ] Trace buffers are cleared at the start of each debug UDF invocation (no cross-call accumulation).

## Out of Scope

- **Real debugger attach** (VS / WinDbg). That's a separate, future escape hatch — Roslyn `EmitOptions` with portable PDBs makes it mechanically possible, but the UX is out of scope here.
- **Tracing inside closures / lambdas** (e.g. inside `.Where(x => ...)`). MVP only traces the top-level statement list and its loops. Phase 2 can extend this.
- **Step-through / breakpoints.** This is batch tracing, not interactive debugging.
- **Rich value rendering** for complex wrappers. MVP renders `ExcelTable` / `ExcelArray` values as a short preview string (`"ExcelTable(7×3)"`); full inline expansion is out of scope.
- **Multi-cell trace aggregation.** `FB.LastTrace()` is single-cell and returns the most recent trace. A `FB.Trace(cellRef)` variant can come later.
- **Persistent trace history.** Buffers live in-process; closing Excel drops them.
- **Tracing across nested FB calls** (a `_DEBUG` UDF that calls another `_DEBUG` UDF). MVP scopes the trace to the outer call only.

