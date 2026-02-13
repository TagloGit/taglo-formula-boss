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
2. User presses `Ctrl+Shift+`` ` (backtick — the Formula Boss shortcut)
3. Floating editor appears, empty for new formulas or with reconstructed source for existing FB formulas
4. User writes formula — no `'` prefix needed, backticks work naturally as DSL delimiters
5. Editor provides syntax highlighting, autocomplete, bracket matching, real-time error indicators
6. User presses Ctrl+Enter or clicks Apply
7. Editor writes quote-prefixed source to cell → normal SheetChange pipeline processes it
8. Editor closes, result appears in cell

**Total extra effort vs. native formula:** One shortcut to open editor.

See [Floating Editor](#floating-editor) for detailed design.

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
6. Cell comment shows: `Error: 'colr' is not a recognised cell property. Did you mean 'color'?`
7. User clicks cell, presses `Ctrl+Shift+`` ` to open floating editor
8. Editor loads the failed formula (still quote-prefixed text in cell), shows `colr` underlined in red
9. User corrects to `color` with help from autocomplete, presses Ctrl+Enter
10. Editor writes corrected source → pipeline processes → formula compiles and works

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

### Journey 7: Row-Wise Aggregation with Column Names

**Scenario:** User needs to sum the product of two columns across all rows.

1. User has a table `tblSales` with columns "Price", "Qty", "Region"
2. User types:
   ```
   '=`tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])`
   ```
3. Add-in detects `tblSales` is an Excel Table, retrieves column names
4. Transpiles to C# that iterates rows with dictionary-based column lookup
5. Result appears immediately

**Column reference syntax:**
- `r[Column Name]` — brackets, spaces allowed in column name
- `r.ColumnName` — dot notation for single-word column names

---

### Journey 8: Robust Column References via LET

**Scenario:** User wants column references that survive table restructuring and benefit from Excel's autocomplete.

1. User has a table `tblCarParks` with columns "Space Start", "Space End", "Zone"
2. User types, using Excel's autocomplete to select header references:
   ```
   '=LET(
       start, tblCarParks[Space Start],
       end, tblCarParks[Space End],
       tbl, tblCarParks,
       `tbl.reduce(0, (acc, r) => acc + r.end - r.start)`
   )
   ```
3. Add-in detects that `r.start` and `r.end` reference LET-bound column variables
4. Transpiles to C# that dynamically looks up columns at runtime
5. Result appears immediately

**Benefits:**
- Excel autocomplete for table/column references
- Refactor-safe: if column names change, update one reference
- Self-documenting: variable names describe the data

---

### Journey 9: Running Totals with State (Scan)

**Scenario:** User needs a running total that resets when a category changes.

1. User has table `tblTransactions` with columns "Category", "Amount"
2. User types:
   ```
   '=LET(
       cat, tblTransactions[Category],
       amt, tblTransactions[Amount],
       tbl, tblTransactions,
       `tbl.scan({sum: 0, lastCat: ""}, (state, r) =>
           LET(reset, r.cat != state.lastCat,
               {sum: IF(reset, r.amt, state.sum + r.amt), lastCat: r.cat})
       ).select(s => s.sum)`
   )
   ```
3. Transpiles to C# with stateful iteration
4. Returns spilling array of running totals

---

### Journey 10: Headerless Data with Index Access

**Scenario:** User has raw data without headers, needs positional column access.

1. User has data in `A1:D100` with no header row
2. User types:
   ```
   '=`A1:D100.reduce(0, (acc, r) => acc + r[0] * r[2])`
   ```
3. Numeric indices `r[0]`, `r[2]` access columns by position (zero-based)
4. Transpiles to direct index access in C#

**Index syntax:**
- `r[0]` — first column
- `r[-1]` — last column (negative index)

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

| Syntax | Meaning | Output Shape |
|--------|---------|--------------|
| `range.cells` | Iterate cells with object model access | 1D (flattened) |
| `range.values` | Iterate values only (fast path) | 1D (flattened) |
| `range.rows` | Iterate as Row objects with column access | 2D (row count changes, columns preserved) |
| `range.cols` | Iterate columns as arrays | 2D (column count changes, rows preserved) |

**Row-wise operations** enable filtering/sorting entire rows with column access:
```
data.rows.where(r => r[0] > 10)           // filter rows where first column > 10
data.rows.orderBy(r => r[Price])          // sort rows by Price column
data.rows.where(r => r.Region == "North") // filter by column name (no spaces)
tbl.reduce(0, (acc, r) => acc + r[Price] * r[Qty])  // aggregate across rows
```

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
| `.select(transform)` | Map/transform elements (flattens to 1D) |
| `.map(transform)` | Transform elements preserving 2D shape |
| `.orderBy(keySelector)` | Sort ascending |
| `.orderByDesc(keySelector)` | Sort descending |
| `.take(n)` | First n elements (negative n takes last n) |
| `.skip(n)` | Skip first n elements (negative n skips last n) |
| `.distinct()` | Remove duplicates |
| `.groupBy(keySelector)` | Group elements |
| `.reduce(seed, func)` | Reduce/fold to single value |
| `.scan(seed, func)` | Running reduction, returns array of intermediate states |
| `.toArray()` | Output as 2D array |
| `.sum()`, `.avg()`, `.min()`, `.max()`, `.count()` | Aggregations |
| `.find(predicate)` | First element matching predicate (or null) |
| `.some(predicate)` | True if any element matches |
| `.every(predicate)` | True if all elements match |

**`.map` vs `.select`:** Use `.map` when you want to transform each cell while keeping the original 2D shape:
```
data.select(v => v * 2)                           // returns 1D array of doubled values
data.map(v => v * 2)                              // returns 2D array same shape as input
data.map(c => c.color == 6 ? c.value * 2 : c.value)  // double yellow cells, preserve shape
```

**Implicit Syntax (convenience features):**

| Feature | Meaning | Example |
|---------|---------|---------|
| Implicit `.values` | Methods called directly on range default to values path | `data.where(v => v > 0)` equals `data.values.where(v => v > 0)` |
| Implicit `.toArray()` | Collection results auto-convert to 2D arrays for Excel | `data.where(v => v > 0)` returns array without explicit `.toArray()` |

---

### Row-Wise Table Operations

Many competitive Excel challenges involve tabular data where you need row-by-row operations — accumulating totals, running calculations, or transforming based on multiple columns per row. The DSL provides natural row-property access that transpiles to efficient C# iteration.

**The Problem with Native Excel:**
```excel
=REDUCE(0, SEQUENCE(ROWS(tbl)), LAMBDA(acc, i,
    LET(row, INDEX(tbl, i, ),
        acc + INDEX(row, 3) * INDEX(row, 5)  -- What are columns 3 and 5?
    )
))
```
The indices are fragile, unreadable, and break if the table structure changes. Native Excel also cannot create reusable helper LAMBDAs that accept callback functions inside REDUCE.

**The DSL Solution:**
```
tbl.reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

#### Row Object and Column Access

When iterating with `.rows`, `.reduce`, `.scan`, or related methods, the lambda parameter is a **Row object** with column access:

| Syntax | Mode | Description |
|--------|------|-------------|
| `r[Column Name]` | Named | Bracket syntax, supports spaces in column names |
| `r.ColumnName` | Named | Dot notation for single-word column names |
| `r[0]`, `r[1]` | Index | Zero-based positional access |
| `r[-1]` | Index | Negative index for last column |

**Named column access** requires headers (see Table Detection below). **Index access** always works.

#### Row-Wise Methods

| Method | Signature | Description |
|--------|-----------|-------------|
| `.reduce(init, fn)` | `(T, (T, Row) => T) => T` | Aggregate rows to single value |
| `.scan(init, fn)` | `(T, (T, Row) => T) => T[]` | Running aggregation, returns array |
| `.where(fn)` | `(Row => bool) => Row[]` | Filter rows by predicate |
| `.select(fn)` | `(Row => U) => U[]` | Transform each row to value |
| `.map(fn)` | `(Row => Row) => Row[]` | Transform rows preserving structure |
| `.find(fn)` | `(Row => bool) => Row?` | First row matching predicate |
| `.some(fn)` | `(Row => bool) => bool` | Any row matches |
| `.every(fn)` | `(Row => bool) => bool` | All rows match |

**Examples:**
```
// Sum product of two columns
tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])

// Filter rows where first column > 10
data.rows.where(r => r[0] > 10)

// Running total that resets on category change
tbl.scan({sum: 0, cat: ""}, (s, r) =>
    LET(reset, r.Category != s.cat,
        {sum: IF(reset, r.Amount, s.sum + r.Amount), cat: r.Category})
).select(s => s.sum)

// Sort rows by third column
data.rows.orderBy(r => r[2])
```

---

### Table Detection and Headers

The DSL automatically detects whether data has headers based on its source:

| Source | Header Detection | Column Access |
|--------|-----------------|---------------|
| Excel Table (`tblName`) | Automatic — uses table headers | Named + Index |
| Range with `.withHeaders()` | First row treated as headers | Named + Index |
| Plain range | No headers | Index only |

**Excel Table Detection:**

Formula Boss detects Excel Tables (ListObjects) by name. When you write:
```
=LET(tbl, tblSales, `tbl.reduce(0, (acc, r) => acc + r[Price] * r[Qty])`)
```

Formula Boss:
1. Recognises `tblSales` as a ListObject in the workbook
2. Retrieves column names from the table headers
3. Enables named column access in row lambdas

**Explicit Headers for Ranges:**

For non-table data with a header row:
```
A1:F100.withHeaders().reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

The `.withHeaders()` modifier tells the DSL to treat the first row as column names.

**Headerless Data:**

For raw data without headers, use index-based access:
```
A1:D100.reduce(0, (acc, r) => acc + r[0] * r[2])
```

#### Row-Wise Transpilation

**C# Output Example:**

DSL input:
```
tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

Generated C#:
```csharp
public static object REDUCEPRICEQTY(object tableRef)
{
    // Get ListObject from table reference
    var listObject = GetListObject(tableRef);
    var data = listObject.DataBodyRange.Value as object[,];
    var headers = BuildHeaderIndex(listObject);  // {"Price": 0, "Qty": 1, ...}

    double acc = 0;
    for (int i = 0; i < data.GetLength(0); i++)
    {
        var price = Convert.ToDouble(data[i, headers["Price"]]);
        var qty = Convert.ToDouble(data[i, headers["Qty"]]);
        acc = acc + price * qty;
    }
    return acc;
}
```

**Robust Mode C# Output:**

DSL input (with LET-bound columns):
```
=LET(price, tblSales[Price], qty, tblSales[Qty], tbl, tblSales,
     `tbl.reduce(0, (acc, r) => acc + r.price * r.qty)`)
```

Generated C#:
```csharp
public static object REDUCEPRICEQTY(object tableRef, string priceCol, string qtyCol)
{
    var listObject = GetListObject(tableRef);
    var data = listObject.DataBodyRange.Value as object[,];
    var headers = BuildHeaderIndex(listObject);

    double acc = 0;
    for (int i = 0; i < data.GetLength(0); i++)
    {
        var price = Convert.ToDouble(data[i, headers[priceCol]]);
        var qty = Convert.ToDouble(data[i, headers[qtyCol]]);
        acc = acc + price * qty;
    }
    return acc;
}
```

The column names are passed as parameters, enabling dynamic lookup.

#### Edge Cases

**Empty table:**
- `.reduce(init, fn)` returns `init`
- `.scan(init, fn)` returns empty array
- `.where(fn)` returns empty array

**Missing column name:**
- Quick mode: Runtime error with clear message ("Column 'Pricee' not found in table. Available columns: Price, Qty, Region")
- Robust mode: Error surfaces when Excel evaluates the invalid column reference

**Numeric column headers:**
- Work with bracket syntax: `r[2024]` for a column named "2024"

**Duplicate column names:**
- First occurrence wins for name-based access
- Use index-based access for subsequent columns

**Mixed header types:**
- Empty header cells: accessible by index only

---

### Deep Property Access

Beyond the shorthand cell properties (`.color`, `.bold`, etc.), users can access the full Excel object model via chained property access.

**Type System:**

The transpiler maintains a type system for Excel COM objects to provide parse-time validation:

| Type | Properties |
|------|------------|
| `Cell` | `Interior`, `Font`, `Value`, `Formula`, `Row`, `Column`, `Address`, etc. |
| `Interior` | `ColorIndex`, `Color`, `Pattern`, `PatternColor`, `PatternColorIndex`, etc. |
| `Font` | `Bold`, `Italic`, `Size`, `Color`, `Name`, `Underline`, etc. |

**Examples:**
```
c.Interior.ColorIndex        // validated: Cell → Interior → ColorIndex
c.Interior.Pattern           // validated: Cell → Interior → Pattern
c.Font.Color                 // validated: Cell → Font → Color
```

**Escape Hatch:**

For properties not in the type system, prefix with `@` to bypass validation:
```
c.@SomeObscureProperty       // passes through verbatim, validated at runtime
c.Interior.@NewExcelProperty // partial validation, then pass-through
```

**Error Messages:**

Invalid properties produce helpful parse-time errors:
```
c.Interior.Patern
// Error: Unknown property 'Patern' on Interior. Did you mean 'Pattern'?
```

---

### Statement Lambdas

For complex logic that can't be expressed as a single expression, statement lambdas allow full C# code blocks.

**Syntax:**
```
data.cells.where(c => {
    var color = (int)(c.Interior.ColorIndex ?? 0);
    var isYellow = color == 6;
    var isOrange = color == 44;
    return isYellow || isOrange;
})
```

**Behaviour:**

- Lexer detects `{` after `=>` and captures the entire block (brace-balanced)
- Block content is emitted as literal C# code
- Lambda parameter (`c`) references work normally
- Type system validation does **not** apply inside statement blocks — Roslyn validates at compile time, COM validates at runtime

**Available Namespaces:**

Statement lambdas have access to these namespaces by default:
```csharp
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
```

**Use Cases:**

- Multi-step calculations with intermediate variables
- Complex conditional logic
- Loops and iteration within a cell's evaluation
- Try/catch for error handling

**Trade-offs:**

| Aspect | Expression Lambda | Statement Lambda |
|--------|-------------------|------------------|
| Validation | Parse-time (type system) | Compile/runtime |
| Error messages | DSL-specific, helpful | Roslyn errors, may be cryptic |
| Complexity | Single expression | Full C# |
| Recommended for | Simple transforms, filters | Complex algorithms |

---

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

**Row-Wise VBA Example:**

DSL input:
```
tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

Generated VBA:
```vba
Function REDUCEPRICEQTY(tbl As Range) As Variant
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To tbl.Columns.Count
        headers(tbl.Cells(1, c).Value) = c
    Next c

    Dim acc As Double
    acc = 0
    Dim i As Long
    For i = 2 To tbl.Rows.Count
        Dim price As Double, qty As Double
        price = tbl.Cells(i, headers("Price")).Value
        qty = tbl.Cells(i, headers("Qty")).Value
        acc = acc + price * qty
    Next i

    REDUCEPRICEQTY = acc
End Function
```

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

## LET Integration and Inline Editing

### Overview

In competitive Excel, most solutions use `LET` functions to structure calculations into named steps. Formula Boss integrates seamlessly with this pattern, enabling:

1. DSL expressions as individual LET steps
2. **Robust column references** that survive table restructuring
3. Dynamic column lookup at runtime (not hardcoded indices)

---

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

---

### Robust Column References

The most powerful LET integration feature is **robust column detection** — referencing table columns in a way that:
- Benefits from Excel's autocomplete
- Survives column reordering or renaming
- Resolves dynamically at runtime

#### User Journey: Robust Mode

1. User has a table `tblSales` with columns "Price", "Qty", "Region"
2. User types, using Excel's autocomplete to select column references:
   ```
   '=LET(
       price, tblSales[Price],
       qty, tblSales[Qty],
       tbl, tblSales,
       `tbl.reduce(0, (acc, r) => acc + r.price * r.qty)`
   )
   ```
3. Formula Boss performs semantic analysis:
   - Detects `tbl` is bound to `tblSales` (an Excel Table)
   - Detects `price` is bound to `tblSales[Price]` (a table column reference)
   - Detects `qty` is bound to `tblSales[Qty]`
   - In the DSL, `r.price` and `r.qty` reference these LET-bound column variables
4. Transpiles to C# that:
   - Accepts column names as string parameters
   - Looks up column indices dynamically from the table at runtime
   - Never hardcodes column positions
5. Result appears immediately

**Benefits:**
- Excel autocomplete for table and column references
- Refactor-safe: if column names change, update one reference
- Self-documenting: variable names describe the data
- Dynamic: UDF always retrieves the current column position from the table

#### Resolution Order for `r.identifier`

When Formula Boss encounters `r.identifier` in a row lambda:

1. **LET-bound column variable** — If `identifier` matches a LET variable bound to a table column reference (e.g., `price, tblSales[Price]`), use that column name dynamically
2. **Numeric literal** — If inside brackets and numeric (e.g., `r[0]`), use as column index
3. **Literal column name** — Otherwise, treat as a literal column name string

#### Semantic Analysis Requirements

Formula Boss parses the surrounding LET structure to build context:

| LET Pattern | Detected As | Effect |
|-------------|-------------|--------|
| `tbl, tblSales` | Table binding | `tbl` refers to the `tblSales` ListObject |
| `price, tblSales[Price]` | Column binding | `price` holds the "Price" column name |
| `data, A1:F20` | Range binding | Plain range, no automatic headers |

**Table Column Reference Parsing:**

The pattern `tableName[ColumnName]` (without `[#Headers]` or `[#All]`) is recognised as a column reference. Formula Boss:
1. Validates `tableName` is a ListObject in the workbook
2. Extracts `ColumnName` as the column identifier
3. Passes the column name as a string parameter to the generated UDF

**Runtime Behaviour:**

The generated UDF does **not** hardcode column indices. Instead:
1. UDF receives table data and column name strings as parameters
2. At runtime, UDF retrieves the actual ListObject to get current column positions
3. Column lookup happens fresh on each calculation

This ensures formulas remain correct even if table columns are reordered.

---

### Self-Documenting Pattern

Each UDF call is preceded by an unused LET variable (prefixed `_src_`) containing the original DSL expression:

```
=LET(
    _src_result, "tbl.reduce(0, (acc, r) => acc + r.price * r.qty)",
    result, RESULT(tbl, price, qty),
    result)
```

This provides:
- **Visibility:** Original expression visible in formula bar
- **Persistence:** Saved with the formula
- **No external dependencies:** No comments or side panels required

---

### Editing Workflow

`Ctrl+Shift+`` ` is the single Formula Boss shortcut. It opens the floating editor, which handles both new formulas and editing existing ones.

**Editing an existing FB formula:**

1. User selects a cell containing a processed Formula Boss LET formula
2. User presses `Ctrl+Shift+`` `
3. Floating editor opens. Formula Boss:
   - Detects `_src_*` variables indicating a processed Formula Boss formula
   - Reconstructs the original formula with backtick expressions
   - Loads the reconstructed source into the editor (without the `'` prefix)
4. Editor shows:
   ```
   =LET(
       price, tblSales[Price],
       qty, tblSales[Qty],
       tbl, tblSales,
       `tbl.reduce(0, (acc, r) => acc + r.price * r.qty)`
   )
   ```
5. User edits the backtick expressions with full editor tooling
6. User presses Ctrl+Enter or clicks Apply
7. Editor writes quote-prefixed source to cell → SheetChange pipeline processes
8. Formula Boss regenerates UDFs (reusing existing UDF names where expressions match)
9. Editor closes

See [Floating Editor](#floating-editor) for the full editor design.

---

### Technical Considerations

**UDF Naming:**
- LET variable name → UDF name (uppercase)
- Collision handling: append numeric suffix if name already exists with different expression

**Chained UDFs:**
- UDFs return values, not cell references
- Second UDF in chain receives array, cannot access cell properties (e.g., `.color`)
- If cell properties needed across multiple logical steps, combine into single expression

**Source Preservation:**
- `_src_` prefix convention for documentation variables
- Variables are evaluated but unused, minimal performance impact

**Edit Mode Implementation:**
- `Application.EnableEvents = false` prevents SheetChange from firing during reconstruction
- `WrapText = false` keeps the multi-line formula readable
- `SendKeys("{F2}")` enters edit mode automatically

**Column Parameter Generation:**

When robust mode detects column references, the generated UDF signature includes string parameters:
```csharp
public static object RESULT(object[,] tbl, string priceCol, string qtyCol)
```

The formula call passes the column reference values:
```
=RESULT(tblSales, tblSales[Price], tblSales[Qty])
```

Excel evaluates `tblSales[Price]` to the column's values, but Formula Boss intercepts this pattern and passes the column name string instead. (Implementation detail: may require wrapping in a helper that extracts the header name.)

---

## Floating Editor

### Overview

The floating editor is a WPF window (AvalonEdit) that provides an assisted editing experience for Formula Boss formulas. It runs on a dedicated STA thread with per-monitor DPI awareness and is positioned centered over the Excel window.

The editor is **not a separate processing path** — it writes the quote-prefixed formula source to the cell on Apply, and the existing SheetChange pipeline handles parsing, compilation, and formula rewriting. This keeps a single code path for all formula processing.

**Key advantages over the formula bar:**

- No `'` prefix needed — the editor adds it on Apply
- Backticks work naturally as DSL delimiters (no Excel syntax conflict)
- Syntax highlighting for both Excel functions and DSL expressions
- Autocomplete for DSL methods, cell properties, column names
- Bracket matching and auto-brace completion
- Real-time parse error indicators before Apply
- Multi-line editing with proper indentation

---

### Activation

`Ctrl+Shift+`` ` (backtick) is the single Formula Boss keyboard shortcut. It is context-aware:

| Cell State | Editor Behaviour |
|------------|-----------------|
| Empty cell | Opens editor empty |
| Cell with unprocessed FB formula (quote-prefixed text) | Opens editor with the formula text (minus `'` prefix) |
| Cell with processed FB LET formula (`_src_` variables) | Reconstructs original source, opens editor with it |
| Cell with processed FB basic formula (no `_src_`) | Opens editor with current cell formula as-is |
| Cell with non-FB formula | Opens editor with current cell formula as-is |

---

### Editor Journeys

#### Journey A: New Formula

1. User selects target cell
2. User presses `Ctrl+Shift+`` `
3. Editor opens empty
4. User writes formula with backtick DSL:
   ```
   =SUM(`data.where(v => v > 0)`)
   ```
5. Editor provides syntax highlighting, autocomplete, real-time error indicators
6. User presses Ctrl+Enter or clicks Apply
7. Editor writes `'=SUM(`data.where(v => v > 0)`)` to the cell (adds `'` prefix)
8. SheetChange fires → existing pipeline parses, compiles, rewrites formula
9. Editor closes

#### Journey B: Edit Existing LET Formula

1. User selects a cell with a processed FB LET formula
2. User presses `Ctrl+Shift+`` `
3. Editor reconstructs original source from `_src_` variables and opens with it
4. User edits DSL expressions with full editor tooling
5. Apply → writes quote-prefixed source → pipeline reprocesses
6. Editor closes

#### Journey C: Edit Existing Basic Formula

1. User selects a cell with a processed FB basic formula (no `_src_`, source not preserved)
2. User presses `Ctrl+Shift+`` `
3. Editor opens with the current cell formula as-is (the rewritten UDF version)
4. User can modify or start fresh

**Note:** Basic formulas currently don't preserve source. Source preservation for basic formulas via auto-wrapping in LET is a potential enhancement — see [User Settings](#user-settings).

#### Journey D: Error Recovery

This is where the editor provides the most value — turning the "edit, Enter, read error comment, repeat" loop into a tight feedback cycle.

1. User types a formula inline (the traditional way), presses Enter
2. Pipeline fails to parse → cell gets error value and comment
3. User presses `Ctrl+Shift+`` ` on the error cell
4. Editor loads the failed formula (still quote-prefixed text in cell)
5. Editor shows parse errors inline (underlines, error messages)
6. User fixes errors with autocomplete and highlighting assistance
7. Apply → pipeline succeeds
8. Editor closes

---

### Editor Features

#### Syntax Highlighting

Custom highlighting definition (`.xshd`) covering:

- DSL methods (`.where`, `.select`, `.reduce`, etc.)
- DSL accessor keywords (`.cells`, `.values`, `.rows`)
- Cell properties (`.color`, `.bold`, `.fontSize`, etc.)
- Strings, numbers, operators
- Lambda arrow (`=>`)

#### Autocomplete

Context-aware completion triggered on `.` or after typing a letter:

| Context | Completion Items |
|---------|-----------------|
| After range/variable `.` | DSL methods and accessors (`.where`, `.cells`, `.reduce`, etc.) |
| After cell parameter `c.` | Cell properties (`.value`, `.color`, `.bold`, etc.) |
| After row parameter `r.` | Column names from the active table (if detectable) |
| General typing | DSL keywords, LET-bound variable names |

Completion accepts on Tab or Enter only. Spacebar and other characters close the completion window (allowing single-letter variable names without interference).

#### Real-Time Parsing

The existing parser is invoked on each keystroke (debounced) to provide:

- Error squiggles on invalid syntax
- Context information for autocomplete (e.g., knowing `c` is a Cell after `.cells.where(c =>`)
- Validation of property names and method chains

This reuses the same parsing infrastructure as the pipeline — not a separate parser.

#### Code Editing Behaviours

- Auto-brace completion: `(`, `[`, `{`, `"`, `` ` `` auto-insert closing pair
- Skip-over: typing a closing character when cursor is before one skips instead of inserting
- Brace de-indentation: `{` snaps to parent indent level
- Block expansion: Enter between `{}` expands to three lines with indentation
- Bracket matching highlight

---

### Range Selection Policy

The editor does **not** support Excel range selection while open. Users should set up range references (in LET variables or directly) before opening the editor. This is the pragmatic approach — Excel's formula bar has deep integration with range selection that cannot be replicated in an external window.

**Workflow:** Set up references first → open editor → write DSL logic → Apply.

---

### Editor ↔ Pipeline Integration

```
Editor (WPF thread)                    Excel thread
       │                                     │
       │  Apply: cell.Formula2 =             │
       │  '=LET(..., `dsl`, ...)  ──────►    │
       │                                     │  SheetChange fires
       │                                     │  FormulaInterceptor detects '= prefix
       │                                     │  FormulaPipeline processes
       │                                     │  UDF compiled and registered
       │                                     │  Cell formula rewritten
       │                                     │
```

The editor writes to the cell via `ExcelAsyncUtil.QueueAsMacro()` to ensure it runs in a valid Excel macro context. The pipeline then handles everything identically to the formula bar path.

---

### Technical Details

- **WPF on separate STA thread** — required for keyboard input when hosted in an Excel add-in
- **Per-monitor DPI awareness** — `SetThreadDpiAwarenessContext(PER_MONITOR_AWARE_V2)` on the WPF thread
- **Window positioning** — Win32 `SetWindowPos` in physical pixels, centered on Excel window
- **HWND created before Show** — `WindowInteropHelper.EnsureHandle()` avoids visible position jump
- **Window style** — `Topmost`, `ToolWindow`, `ShowInTaskbar=True` for Alt+Tab visibility
- **Error isolation** — all editor event handlers wrapped in try-catch to prevent crashing Excel

---

## User Settings

A collection of user-configurable preferences. Settings are persisted to the user's AppData folder.

| Setting | Default | Description |
|---------|---------|-------------|
| Editor shortcut | `Ctrl+Shift+`` ` | Keyboard shortcut to open the floating editor |
| Auto-wrap basic formulas in LET | Off | When enabled, basic (non-LET) FB formulas are automatically wrapped in a LET structure to preserve the DSL source via `_src_` variables. This enables full round-trip editing of all FB formulas, at the cost of changing the cell formula shape. |

*This section will grow as new configurable behaviours are identified.*
