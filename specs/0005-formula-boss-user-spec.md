# 0005 — Formula Boss User Specification

What users can write and what results to expect. This is the source of truth for AddIn tests.

---

## Overview

Formula Boss is an Excel add-in that lets power users write inline C# expressions inside Excel formulas. Expressions operate on typed wrappers around Excel data and compile to UDFs at runtime. The primary use case is competitive Excel, where speed and expressiveness matter.

---

## Entry Points

### 1. Quote-Prefix (Fast Path)

Type a formula with backtick-delimited expressions directly into a cell:

```
'=SUM(`data.Cells.Where(c => c.Color == 6).Select(c => c.Value)`)
```

1. Prefix with `'` so Excel treats it as text
2. Press Enter
3. Add-in detects the pattern, compiles the backtick expression to a UDF
4. Cell formula is rewritten to call the generated UDF (e.g. `=SUM(__udf_a1b2(data))`)
5. Result appears

Multiple backtick expressions in one formula are supported:

```
'=`A1:C10.Sum()` + `D1:D10.Count()`
```

### 2. Floating Editor (Assisted Path)

1. Select target cell
2. Press `Ctrl+Shift+`` ` (backtick)
3. Floating editor opens with syntax highlighting, bracket matching, error squiggles, and Roslyn-powered autocomplete
4. Write expression (multi-line supported)
5. Press `Ctrl+Enter` to apply

### 3. LET Integration

Backtick expressions work inside LET formulas. LET-bound variables are automatically captured as UDF parameters:

```
=LET(
    maxPop, XLOOKUP(...),
    pConts, TEXTSPLIT(...),
    result, `tblCountries.Rows.Where(r =>
        r["Population"] < maxPop
        && pConts.Any(c => c == r["Continent"]))`,
    result)
```

`tblCountries`, `maxPop`, and `pConts` are all detected as free variables. Each LET binding with backticks is processed independently.

---

## Expression Language

Expressions are **standard C#**. All external identifiers (not lambda parameters, not C# keywords, not local variables) are automatically detected and become UDF parameters.

### Parameter Detection

```
`tblCountries.Rows.Where(r => r["Population"] > maxPop)`
```

- `tblCountries` → free variable → UDF parameter (Excel resolves as table/range/named range)
- `maxPop` → free variable → UDF parameter
- `r` → lambda parameter → not a UDF parameter
- `Where`, `Rows` → method/property names → not parameters

No explicit parameter syntax is needed. All parameters are equal — there is no "primary input" concept.

### Range References

Excel-style range references are supported directly in expressions:

```
`A1:C10.Sum()`
`A1:C10.Rows.Where(r => r[0] > B1)`
```

Range references like `A1:C10` are automatically converted to valid C# identifiers internally. Single-cell references like `A1` or `B1` work as free variables — Excel resolves them as cell references.

**Cross-sheet references** (e.g. `Sheet2!A1:B10`) are not currently supported directly in expression bodies. Use LET binding as a workaround:

```
=LET(data, Sheet2!A1:B10, `data.Sum()`)
```

> **Planned:** Cross-sheet range reference support in expression bodies (#100).

### Statement Blocks

For complex logic, wrap the expression in braces:

```
`{
    var count = tbl.Rows.Count();
    if (count > threshold)
        return tbl.Rows.Where(r => r["Price"] > 5).ToRange();
    return otherTable.Sum();
}`
```

Statement blocks support `var` declarations, `if`/`else`, `for`/`foreach`, and `return`. All external identifiers are still detected as free variables.

---

## Type System

Every UDF parameter is wrapped by `ExcelValue.Wrap()`, which inspects the runtime value and returns the appropriate type:

| Input | Wrapped Type | When |
|---|---|---|
| Excel Table (ListObject) with headers detected | `ExcelTable` | Expression uses `r["ColumnName"]` bracket access |
| `object[,]` (range, TEXTSPLIT, SORT, etc.) | `ExcelArray` | Multi-cell input without header access |
| Single value | `ExcelScalar` | Scalar inputs (single cells, LET scalars) |

All three types extend `ExcelValue`, which supports the full operation API. No casting is ever needed.

### ExcelScalar — Single Value

Wraps one value. All collection operations work with single-element semantics:

```
`myCell.Sum()`       // returns the value itself
`myCell.Count()`     // returns 1
`myCell > 5`         // comparison via operator overloading
`myCell * 2 + 1`     // arithmetic via operator overloading
```

### ExcelArray — 2D Array

Wraps `object[,]`. Always 2D internally (even 1×N or N×1).

```
`myRange.Sum()`                          // sum all elements
`myRange.Where(x => x > 0)`             // element-wise filter
`myRange.Select(x => x * 2)`            // element-wise project (flattens to 1 column)
`myRange.Map(x => x * 2)`               // element-wise transform (preserves 2D shape)
`myRange.Rows.Where(r => r[0] > 10)`    // row-wise filter
```

### ExcelTable — Array with Column Names

Extends `ExcelArray` with column metadata. Created when header access (`r["Col"]`) is detected:

```
`tbl.Rows.Where(r => r["Price"] > 100)`
`tbl.Rows.OrderBy(r => r["Name"])`
`tbl.Rows.Select(r => r["Price"] * r["Qty"])`
```

Headers are extracted from the first row of the range. When using Excel Tables (ListObjects), the `[#All]` modifier is automatically appended so headers are included.

---

## Operations

### Element-Wise Operations (on ExcelValue / IExcelRange)

These iterate over all values in row-major order (left-to-right, top-to-bottom). The lambda parameter is `ExcelValue`:

| Operation | Returns | Description |
|---|---|---|
| `.Where(x => pred)` | `IExcelRange` | Filter elements |
| `.Select(x => expr)` | `IExcelRange` | Project elements (flattens to 1 column) |
| `.SelectMany(x => seq)` | `IExcelRange` | Flatten nested sequences |
| `.Map(x => expr)` | `IExcelRange` | Transform elements (preserves 2D shape) |
| `.Any(x => pred)` | `bool` | True if any element matches |
| `.All(x => pred)` | `bool` | True if all elements match |
| `.First(x => pred)` | `ExcelValue` | First matching element |
| `.FirstOrDefault(x => pred)` | `ExcelValue?` | First match or null |
| `.OrderBy(x => key)` | `IExcelRange` | Sort ascending |
| `.OrderByDescending(x => key)` | `IExcelRange` | Sort descending |
| `.Take(n)` | `IExcelRange` | First n elements (negative = last n) |
| `.Skip(n)` | `IExcelRange` | Skip first n (negative = skip last n) |
| `.Distinct()` | `IExcelRange` | Remove duplicates |
| `.Aggregate(seed, (acc, x) => expr)` | `ExcelValue` | Fold to single value |
| `.Scan(seed, (acc, x) => expr)` | `IExcelRange` | Running fold (returns all intermediate values) |
| `.Count()` | `int` | Number of elements |
| `.Sum()` | `ExcelScalar` | Sum of numeric values |
| `.Min()` | `ExcelScalar` | Minimum value |
| `.Max()` | `ExcelScalar` | Maximum value |
| `.Average()` | `ExcelScalar` | Mean of numeric values |

### Row-Wise Operations (on RowCollection via `.Rows`)

Access `.Rows` to iterate row-by-row. The lambda parameter is `dynamic` (a `Row` object), enabling column access:

| Operation | Returns | Description |
|---|---|---|
| `.Rows.Where(r => pred)` | `RowCollection` | Filter rows |
| `.Rows.Select(r => expr)` | `IExcelRange` | Project each row to a value |
| `.Rows.Any(r => pred)` | `bool` | True if any row matches |
| `.Rows.All(r => pred)` | `bool` | True if all rows match |
| `.Rows.First(r => pred)` | `Row` | First matching row |
| `.Rows.FirstOrDefault(r => pred)` | `Row?` | First match or null |
| `.Rows.OrderBy(r => key)` | `RowCollection` | Sort rows ascending |
| `.Rows.OrderByDescending(r => key)` | `RowCollection` | Sort rows descending |
| `.Rows.Take(n)` | `RowCollection` | First n rows (negative = last n) |
| `.Rows.Skip(n)` | `RowCollection` | Skip first n rows (negative = skip last n) |
| `.Rows.Distinct()` | `RowCollection` | Remove duplicate rows |
| `.Rows.Count()` | `int` | Number of rows |
| `.Rows.ToRange()` | `IExcelRange` | Convert back to array for element-wise ops |

> **Planned — not yet implemented:**
>
> | Operation | Returns | Description | Issue |
> |---|---|---|---|
> | `.Rows.Aggregate(seed, (acc, r) => expr)` | value | Fold rows to single value | #99 |
> | `.Rows.Scan(seed, (acc, r) => expr)` | `RowCollection` | Running fold over rows | #99 |
> | `.Rows.GroupBy(r => key)` | TBD | Group rows by key (API design pending) | #101 |

### Column Access on Rows

Inside row lambdas, access columns via:

| Syntax | Example | Notes |
|---|---|---|
| `r["Column Name"]` | `r["Price"]` | String bracket access — always works |
| `r[0]`, `r[-1]` | `r[0]` (first), `r[-1]` (last) | Numeric index — zero-based, negative supported |
| `r.ColumnName` | `r.Price` | Dot notation — intellisense supported, rewritten to bracket access before compilation |

**Dot notation details:** When table headers are known, the editor provides intellisense completions for column names. `r.Population2025` is rewritten to `r["Population 2025"]` before compilation. If two columns sanitise to the same identifier (e.g. `Foo Bar` and `FooBar`), those columns fall back to bracket-only access.

Column access returns a `ColumnValue` which supports:
- Comparison operators: `r["Price"] > 100`, `r["Name"] == "USA"`
- Arithmetic operators: `r["Price"] * r["Qty"]`
- Implicit conversion to `double`, `string`, `bool`
- Cross-type comparison with `ExcelValue`: `r["Price"] > maxPrice` (where `maxPrice` is a scalar parameter)

### Cell Formatting Access

Access cell formatting (color, bold, font) via `.Cells` or `.Cell`:

```
// Filter by cell color
`data.Cells.Where(c => c.Color == 6).Select(c => c.Value)`

// Access formatting via row column
`tbl.Rows.Where(r => r["Status"].Cell.Bold)`

// Sum cells with specific background color
`data.Cells.Where(c => c.Rgb == 255).Sum()`
```

Cell properties:

| Property | Type | Description |
|---|---|---|
| `c.Value` | `object` | Cell value |
| `c.Formula` | `string` | Cell formula text |
| `c.Format` | `string` | Number format string |
| `c.Address` | `string` | Sheet-qualified address (e.g. `Sheet1!A1`) |
| `c.Row` | `int` | Row number (1-based) |
| `c.Col` | `int` | Column number (1-based) |
| `c.Color` | `int` | Interior ColorIndex |
| `c.Rgb` | `int` | Interior RGB color |
| `c.Bold` | `bool` | Font bold |
| `c.Italic` | `bool` | Font italic |
| `c.FontSize` | `double` | Font size |
| `c.Interior` | `Interior` | Sub-object: `.ColorIndex`, `.Color` |
| `c.Font` | `CellFont` | Sub-object: `.Bold`, `.Italic`, `.Size`, `.Name`, `.Color` |

Cell access requires COM interop and is automatically detected — UDFs using `.Cell` or `.Cells` are registered with `IsMacroType = true`.

**Aggregation extensions** on `IEnumerable<Cell>`:

```
`data.Cells.Where(c => c.Color == 6).Sum()`
`data.Cells.Where(c => c.Bold).Count()`
`data.Cells.Where(c => c.Color == 6).Average()`
`data.Cells.Min()`
`data.Cells.Max()`
```

---

## Intellisense

The floating editor provides Roslyn-powered autocomplete:

- **Wrapper type methods:** `.Rows`, `.Where()`, `.Any()`, `.Sum()`, etc.
- **Column names:** After `r.` or `r["`, column names from table headers are suggested
- **Standard C# methods:** String methods (`.Contains()`, `.Split()`), Math, `Regex`, etc.
- **LINQ methods:** Available on any `IEnumerable<>`
- **Named ranges and table names:** Suggested as top-level identifiers
- **Real-time error squiggles:** Syntax errors highlighted as you type

---

## Result Handling

| Expression Result | Excel Display |
|---|---|
| Single value (number, string, bool) | Displayed in the formula cell |
| Multi-cell result (`IExcelRange`, `RowCollection`, etc.) | Spills as a dynamic array via `Formula2` |
| `Row` | Spills as a single row |
| `ColumnValue` | Displayed as single value |

---

## Error Handling

- **Compile errors:** Displayed as a cell comment on the formula cell. The cell retains its backtick text.
- **Runtime errors:** Currently surface as `#VALUE!` from Excel. No structured runtime error reporting yet.

---

## Examples

### Sum yellow cells
```
'=SUM(`data.Cells.Where(c => c.Color == 6).Select(c => c.Value)`)
```

### Filter table rows
```
'=`tblSales.Rows.Where(r => r["Region"] == "EMEA" && r["Revenue"] > 1000)`
```

### Multi-input expression
```
=LET(
    maxPop, XLOOKUP(...),
    continents, TEXTSPLIT("Europe,Asia", ","),
    result, `tblCountries.Rows.Where(r =>
        r["Population"] < maxPop
        && continents.Any(c => c == r["Continent"]))`,
    result)
```

### Top 5 by descending price
```
'=`tbl.Rows.OrderByDescending(r => r["Price"]).Take(5)`
```

### Last 3 elements of a range
```
'=`A1:A20.Take(-3)`
```

### Element-wise transform preserving shape
```
'=`myRange.Map(x => x * 1.1)`
```

### Conditional with statement block
```
'=`{
    var total = sales.Sum();
    if (total > target)
        return "Over target by " + (total - target);
    return "Under target by " + (target - total);
}`
```

### Cross-type comparison
```
=LET(
    threshold, B1,
    result, `A1:A100.Where(x => x > threshold)`,
    result)
```
