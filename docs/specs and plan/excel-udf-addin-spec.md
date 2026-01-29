# Excel Inline UDF Add-in Specification

## Overview

An Excel add-in that allows power users to write inline expressions using a concise DSL (domain-specific language) that transpile to C# UDFs at runtime. The primary use case is competitive Excel, where speed of development matters and challenges often require capabilities beyond native Excel formulas.

---

## Core Value Proposition

- Write terse, expressive code inline within Excel formulas
- Access Excel object model properties (cell colour, formatting) that are impossible with native formulas
- Use familiar programming constructs (LINQ-style operations, proper iteration, recursion)
- Minimal context switching — stay in the flow of formula editing

---

## Dream User Journeys

### Journey 1: Inline Expression — Fast Path

**Scenario:** User needs to sum all yellow-highlighted cells in a range.

1. User selects target cell for result
2. User types: `'=SUM(`data.cells.where(c=>c.color==6).values`)`
3. User presses Enter
4. Formula briefly appears as text
5. Add-in detects the pattern, parses the backtick expression
6. Add-in transpiles to C#, compiles, registers UDF (sub-second)
7. Cell formula is rewritten to: `=SUM(__udf_a1b2(data))`
8. Result appears immediately

**Total extra effort vs. native formula:** One apostrophe, backtick syntax.

---

### Journey 2: Floating Editor — Assisted Path

**Scenario:** User needs to write a complex multi-step transformation and wants tooling assistance.

1. User selects target cell
2. User presses `Ctrl+Shift+E` (not in edit mode)
3. Floating editor appears, pre-loaded with existing formula if any
4. Editor provides:
   - Syntax highlighting for Excel functions and DSL
   - Autocomplete for DSL keywords and methods
   - Real-time error indicators
   - Signature help showing available cell properties
5. User types formula with backtick expressions
6. User presses Enter or clicks Apply
7. Add-in processes, compiles, rewrites formula
8. Editor closes, result appears in cell

**Total extra effort vs. native formula:** Shortcut to open editor, otherwise similar.

---

### Journey 3: Reusable Named UDF

**Scenario:** User realises they'll need the same transformation multiple times.

1. User writes inline expression as in Journey 1 or 2
2. User selects the cell, presses `Ctrl+Shift+N` (name UDF)
3. Dialog prompts for function name: user enters `SumYellow`
4. Add-in registers the UDF permanently (for this session) under that name
5. User can now use `=SumYellow(A1:J10)` anywhere
6. Optionally: add-in offers to save to a personal library file for future sessions

---

### Journey 4: Debugging a Failed Expression

**Scenario:** User writes an expression with a typo or logic error.

1. User types `'=LET(x, `data.cells.where(c=>c.colr==6).values`, SUM(x))`
2. User presses Enter
3. Add-in detects backtick expression, attempts to parse
4. Parse/compile fails: `colr` is not a valid property
5. Cell displays: `#UDF_ERR`
6. Add-in task pane (or tooltip on hover) shows: `Error: 'colr' is not a recognised cell property. Did you mean 'color'?`
7. User clicks cell, presses `Ctrl+Shift+E` to open editor
8. Editor shows formula with `colr` underlined in red
9. User corrects to `color`, presses Enter
10. Formula compiles and works

---

### Journey 5: Pure Value Transformation (No Object Model)

**Scenario:** User needs to filter and transform values using LINQ-style operations, but doesn't need object model access.

1. User types: `'=`nums.where(n=>n>0).select(n=>n*2).toArray()`
2. Add-in recognises this as value-only operation (no `.cells`, no `.color` etc.)
3. Transpiles to fast path — no COM interop, pure array manipulation
4. Executes with minimal overhead

---

### Journey 6: Complex Algorithm — Graph Shortest Path

**Scenario:** Challenge requires finding shortest path in a network defined by an edge list.

1. User has edge list in `A1:C20` (from, to, weight)
2. User types in editor:
   ```
   =`edges.shortestPath(from: col(0), to: col(1), weight: col(2), start: "A", end: "Z")`
   ```
3. Add-in recognises `shortestPath` as a built-in algorithm in the standard library
4. Transpiles to pre-written Dijkstra implementation
5. Returns path length (or full path as spill array)

---

## MVP User Journeys

### MVP Journey 1: Quote-Prefix Inline Expression

**Minimum implementation of the fast path.**

1. User types: `'=SUM(`A1:J10.where(c=>c.color==6).values`)`
2. User presses Enter
3. Worksheet change handler detects cell starting with `'=` containing backticks
4. Extracts backtick expression
5. Transpiles using simple template substitution
6. Compiles with Roslyn
7. Registers UDF with ExcelDNA
8. Rewrites cell formula
9. If error: cell shows `#UDF_ERR`, error message in adjacent comment or task pane

**Supported DSL for MVP:**

- `range.cells` — iterate with object model access (required for color, formatting, etc.)
- `range.values` — iterate values only (fast path) — **implicit if omitted**
- `.where(c => condition)` — filter
- `.select(c => expression)` — map
- `.toArray()` — materialise to 2D output — **implicit for collection results**
- Cell properties: `.value`, `.color`, `.row`, `.col`
- Basic operators: `==`, `!=`, `>`, `<`, `>=`, `<=`, `&&`, `||`, `+`, `-`, `*`, `/`

**Simplified syntax examples:**
- `data.where(v => v > 0)` — same as `data.values.where(v => v > 0).toArray()`
- `data.cells.where(c => c.color == 6).select(c => c.value)` — cell filtering with implicit `.toArray()`

**Not in MVP:**

- Floating editor
- Named UDF persistence
- Autocomplete/intellisense
- Built-in algorithm library

---

### MVP Journey 2: Basic Error Feedback

1. User types malformed expression
2. Add-in attempts parse, fails
3. Cell value set to `#UDF_ERR`
4. Cell comment added with error message
5. User can read comment, correct formula, retry

---

### MVP Journey 3: Range vs. Values Automatic Detection

1. If expression uses `.cells` or object model properties (`.color`, `.bold`, etc.), use range reference path
2. If expression uses only `.values` and value operations, use fast value array path
3. User doesn't need to think about this — add-in infers from expression content

---

## Non-UI Capabilities

### DSL Feature Set

**Cell/Range Access:**

| Syntax | Meaning |
|--------|---------|
| `range.cells` | Iterate cells with object model access |
| `range.values` | Iterate values only (fast path) |
| `range.rows` | Iterate rows as arrays |
| `range.cols` | Iterate columns as arrays |

**Cell Properties (object model required):**

| Property | Type | Notes |
|----------|------|-------|
| `.value` | variant | Cell value |
| `.color` | int | Interior.ColorIndex |
| `.rgb` | int | Interior.Color (RGB) |
| `.bold` | bool | Font.Bold |
| `.italic` | bool | Font.Italic |
| `.fontSize` | int | Font.Size |
| `.format` | string | NumberFormat |
| `.formula` | string | Cell formula |
| `.row` | int | Row number |
| `.col` | int | Column number |
| `.address` | string | Cell address |

**LINQ-Style Operations:**

| Method | Description |
|--------|-------------|
| `.where(predicate)` | Filter elements |
| `.select(transform)` | Map/transform elements |
| `.orderBy(keySelector)` | Sort ascending |
| `.orderByDesc(keySelector)` | Sort descending |
| `.take(n)` | First n elements |
| `.skip(n)` | Skip first n elements |
| `.distinct()` | Remove duplicates |
| `.groupBy(keySelector)` | Group elements |
| `.aggregate(seed, func)` | Reduce/fold |
| `.toArray()` | Output as 2D array |
| `.sum()`, `.avg()`, `.min()`, `.max()`, `.count()` | Aggregations |

**Implicit Syntax (convenience features):**

| Feature | Meaning | Example |
|---------|---------|---------|
| Implicit `.values` | Methods called directly on range default to values path | `data.where(v => v > 0)` equals `data.values.where(v => v > 0)` |
| Implicit `.toArray()` | Collection results auto-convert to 2D arrays for Excel | `data.where(v => v > 0)` returns array without explicit `.toArray()` |

**Built-in Algorithms (post-MVP):**

| Function | Description |
|----------|-------------|
| `.shortestPath(...)` | Dijkstra on edge list |
| `.connectedComponents()` | Graph connectivity |
| `.topoSort()` | Topological ordering |
| `.permutations()` | Generate permutations |
| `.combinations(k)` | Generate k-combinations |
| `.subsets()` | Power set |
| `.iterate(func, untilCondition)` | Recursive iteration |

---

### Transpilation Behaviour

**Input types:**

- Range reference → passed as `ExcelReference`, converted to COM `Range` if object model needed
- Value array → passed as `object[,]`, processed directly
- Scalar → passed as `object`, handled appropriately

**Output types:**

- All outputs normalised to `object[,]` for spill compatibility
- Scalars wrapped in 1x1 array
- Ragged results padded with `null` (appears blank)

**Caching:**

- Identical backtick expressions generate identical UDF names (hash-based)
- Re-entering same expression reuses existing compiled UDF
- No recompilation for repeated patterns

---

### Error Handling

**Parse errors:**

- Invalid syntax in backtick expression
- Unknown property or method name
- Type mismatches in expression

**Compile errors:**

- Generated C# fails to compile (indicates transpiler bug)

**Runtime errors:**

- Object model access fails (e.g., merged cells, protected sheet)
- Type coercion fails
- Algorithm-specific errors (e.g., graph has no path)

**Error reporting:**

- Cell displays `#UDF_ERR` or `#UDF_PARSE` or `#UDF_RUNTIME`
- Detailed message available via cell comment, task pane, or hover tooltip

---

### Session and Persistence

**Session-scoped (default):**

- Compiled UDFs live in memory for current Excel session
- Lost on Excel close

**Persisted (optional, post-MVP):**

- User can save UDF to personal library
- Library loaded on add-in startup
- Stored as JSON/XML mapping expression → compiled DLL or source

---

### Performance Characteristics

| Operation | Expected Performance |
|-----------|---------------------|
| First Roslyn compile | 1-2 seconds (one-time load) |
| Subsequent compiles | 50-200ms |
| Value-only operations | Near-native speed |
| Object model iteration | ~1000 cells/second |
| Formula rewrite | <50ms |

**Guidance for users:**

- Prefer `.values` over `.cells` when object model not needed
- Keep object-model ranges under 10,000 cells for interactive use
- Large datasets: consider helper columns to pre-filter

---

## Export and Portability

### The Problem

During competition, UDFs run via ExcelDNA. When the workbook is submitted:

- Judges don't have the add-in installed
- Formulas referencing generated UDFs return `#NAME?`
- Workbook is broken

### Solution: Dual Transpilation

The DSL transpiles to two backends:

```
DSL Expression
     │
     ├──► C# Transpiler ──► Roslyn ──► ExcelDNA UDF (fast, competition use)
     │
     └──► VBA Transpiler ──► Workbook Module (portable, export use)
```

**During competition:** C# backend for maximum execution and development speed.

**On export:** VBA backend for portability. Generated VBA can be verbose — correctness is all that matters.

### Export Workflow

1. User completes challenge
2. User clicks "Prepare for Export" (ribbon button or shortcut)
3. Add-in identifies all cells using generated UDFs
4. For each unique UDF:
   - Transpiles DSL expression to equivalent VBA function
   - Injects into workbook's VBA project
5. Formulas remain unchanged — they now reference VBA functions instead of ExcelDNA
6. Workbook saved as `.xlsm`

### Export Options

| Option | Behaviour |
|--------|-----------|
| **Transpile to VBA** (default) | Formulas preserved, workbook self-contained |
| **Bake to values** | Formulas replaced with static values, smallest file |
| **Both** | VBA injected, but a backup sheet stores baked values |

### VBA Transpilation Notes

- LINQ-style operations become explicit loops
- Object model access translates directly (VBA is native to Excel)
- Generated code prioritises correctness over elegance
- Function names match ExcelDNA names for seamless switchover

### Trust Settings Requirement

Programmatic VBA injection requires:

- Excel setting: "Trust access to the VBA project object model"
- Located in: File → Options → Trust Center → Trust Center Settings → Macro Settings
- One-time setup per machine

### Edge Cases

- **No UDFs used:** Export does nothing, workbook already portable
- **Mixed UDFs and native formulas:** Only UDF cells affected
- **Circular references involving UDFs:** VBA equivalent maintains same dependency structure
- **Named UDFs:** Exported with user-assigned names

---

## Open Questions

1. **Syntax for multi-expression pipelines:** Should we allow multiple backtick regions in one formula, or require a single expression?

2. **Lambda syntax:** Use `c => c.value` (C# style) or `c -> c.value` (more universal) or something terser?

3. **Naming convention for generated UDFs:** Hash-based (`__udf_a1b2c3`) or descriptive (`__udf_colorfilter_6`)?

4. **Conflict handling:** What if user's named UDF conflicts with future built-in? Namespace prefix?

5. **Array orientation:** Default to column output, row output, or infer from context?

6. **Error granularity:** One error type or multiple (`#UDF_ERR`, `#UDF_PARSE`, `#UDF_TYPE`, etc.)?

---

## Future Possibilities

- **Collaborative library:** Share UDFs/expressions with other competitive Excel users
- **Expression history:** Recall and reuse recent expressions
- **Performance profiler:** Show timing breakdown for complex expressions
- **Test mode:** Validate expression against expected output before committing
- **Export to VBA:** Generate standalone VBA for environments where add-in isn't available

---

## Future Feature: LET Integration and Inline Editing

### Overview

In competitive Excel, most solutions use `LET` functions to structure calculations into named steps. Formula Boss should integrate seamlessly with this pattern, allowing users to add DSL expressions as individual LET steps.

### User Journey: LET Step Integration

1. User writes a LET formula with backtick expressions for specific steps:
   ```
   '=LET(data, A1:F20,
        coloredCells, `data.cells.where(c => c.color != -4142)`,
        result, `coloredCells.select(c => c.value * 2).toArray()`,
        result)
   ```

2. Formula Boss processes each backtick expression:
   - Names the UDF based on the LET variable name (e.g., `COLOREDCELLS`, `RESULT`)
   - Preserves the original DSL as a documentation step

3. Final formula becomes:
   ```
   =LET(data, A1:F20,
        _src_coloredCells, "data.cells.where(c => c.color != -4142)",
        coloredCells, COLOREDCELLS(data),
        _src_result, "coloredCells.select(c => c.value * 2).toArray()",
        result, RESULT(coloredCells),
        result)
   ```

### Self-Documenting Pattern

Each UDF call is preceded by an unused LET variable (prefixed `_src_`) containing the original DSL expression. This provides:

- **Visibility:** Original expression visible in formula bar
- **Persistence:** Saved with the formula
- **No external dependencies:** No comments or side panels required

### Editing Workflow

1. User wants to edit an existing Formula Boss LET formula
2. User presses editing shortcut (e.g., `Ctrl+Shift+E`)
3. Formula Boss:
   - Reads the `_src_*` variables
   - Reconstructs the original formula with backtick expressions
   - Adds quote prefix to enter edit mode
4. Cell shows:
   ```
   '=LET(data, A1:F20,
        coloredCells, `data.cells.where(c => c.color != -4142)`,
        result, `coloredCells.select(c => c.value * 2).toArray()`,
        result)
   ```
5. User edits the backtick expressions
6. User presses Enter
7. Formula Boss regenerates UDFs (overwriting existing UDFs with same names)

### Technical Considerations

**UDF Naming:**
- LET variable name → UDF name (uppercase)
- Collision handling: append hash suffix if name already exists with different expression

**Chained UDFs:**
- UDFs return values, not cell references
- Second UDF in chain receives array, cannot access cell properties (e.g., `.color`)
- If cell properties needed across multiple logical steps, combine into single expression

**Source Preservation:**
- `_src_` prefix convention for documentation variables
- Variables are evaluated but unused, minimal performance impact
