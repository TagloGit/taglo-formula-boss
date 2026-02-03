## Row-Wise Table Operations

### The Problem

Many competitive Excel challenges involve tabular data where you need to perform row-by-row operations — accumulating totals, running calculations, or transforming data based on multiple columns per row. Native Excel handles this with REDUCE or SCAN, but accessing columns by name within the iteration is painful:

```excel
=REDUCE(0, SEQUENCE(ROWS(tbl)), LAMBDA(acc, i,
    LET(
        row, INDEX(tbl, i, ),
        acc + INDEX(row, 3) * INDEX(row, 5)  -- What are columns 3 and 5?
    )
))
```

The indices are fragile, unreadable, and break if the table structure changes.

**Native Excel limitation:** You cannot create a reusable helper LAMBDA in Name Manager that accepts a callback function and calls it inside REDUCE. Excel's scoping rules prevent this pattern entirely, forcing users to copy-paste boilerplate for every row-wise operation.

### The Solution

The DSL provides natural row-property access that transpiles to efficient C# iteration:

```
tbl.reduce(0, (acc, row) => acc + row.Price * row.Qty)
```

The transpiler handles column lookups automatically.

---

### User Journey 7: Row-Wise Aggregation — Quick Mode

**Scenario:** User needs to sum the product of two columns across all rows.

1. User has a table `tblSales` with columns "Price", "Qty", "Region"
2. User types:
   ```
   '=`tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])`
   ```
3. Add-in parses the expression, identifies column references `r[Price]` and `r[Qty]`
4. Transpiles to C# that iterates rows with dictionary-based column lookup
5. Result appears immediately

**Column reference syntax (Quick Mode):**
- `r[Column Name]` — brackets without quotes, spaces allowed
- `r.ColumnName` — dot notation for single-word column names (no spaces)

**Total extra effort:** Backtick wrapper, bracket syntax for column names.

---

### User Journey 8: Row-Wise Aggregation — Robust Mode

**Scenario:** User wants column references that survive table restructuring and benefit from Excel's autocomplete.

1. User has a table `tblCarParks` with columns "Space Start", "Space End", "Zone"
2. User types, using Excel's autocomplete and arrow keys to select header references:
   ```
   '=LET(
       start, tblCarParks[[#Headers],[Space Start]],
       end, tblCarParks[[#Headers],[Space End]],
       tbl, tblCarParks[#All],
       `tbl.reduce(0, (acc, r) => acc + r.end - r.start)`
   )
   ```
3. Add-in detects that `r.start` and `r.end` reference LET-bound variables containing column names
4. Transpiles to C# using those variable values for column lookup
5. Result appears immediately

**Benefits of Robust Mode:**
- Excel autocomplete for table/column references
- Arrow key navigation to select headers
- Refactor-safe: if column names change, update one reference
- Self-documenting: variable names describe the data

---

### User Journey 9: Row-Wise Scan (Running Totals)

**Scenario:** User needs a running total that resets when a category changes.

1. User has table `tblTransactions` with columns "Category", "Amount"
2. User types:
   ```
   '=LET(
       cat, tblTransactions[[#Headers],[Category]],
       amt, tblTransactions[[#Headers],[Amount]],
       tbl, tblTransactions[#All],
       `tbl.scan({sum: 0, lastCat: ""}, (state, r) => 
           LET(reset, r.cat != state.lastCat,
               {sum: IF(reset, r.amt, state.sum + r.amt), lastCat: r.cat})
       ).select(s => s.sum)`
   )
   ```
3. Transpiles to C# with stateful iteration
4. Returns spilling array of running totals

---

### User Journey 10: Headerless Data — Index Mode

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
- `r[1]` — second column
- Negative indices: `r[-1]` — last column

---

### Data Source Types

The row-wise operations work with multiple data source types:

| Source | Syntax | Notes |
|--------|--------|-------|
| Excel Table (with headers) | `tblName` or `tblName[#All]` | Full structured reference support |
| Range with header row | `A1:F100.withHeaders()` | First row treated as headers |
| Range without headers | `A1:F100` | Index-based access only |
| LET-bound range | `tbl` (from `LET(tbl, ...)`) | Inherits type from source |

**Automatic detection:**
- If source is an Excel Table object → headers available
- If `.withHeaders()` is called → first row becomes headers
- Otherwise → index-based access only

---

### Column Reference Syntax Summary

| Syntax | Mode | Use Case |
|--------|------|----------|
| `r[Column Name]` | Quick | Hardcoded column name with spaces |
| `r.ColumnName` | Quick | Hardcoded single-word column name |
| `r.varName` | Robust | Variable bound to header reference in outer LET |
| `r[0]`, `r[1]` | Index | Positional access, headerless data |
| `r[-1]` | Index | Last column (negative index) |

**Resolution order:**
1. If identifier matches a LET-bound variable → use that variable's value as column name
2. Else if numeric → use as column index
3. Else → treat as literal column name

---

### DSL Methods for Row-Wise Operations

| Method | Signature | Description |
|--------|-----------|-------------|
| `.reduce(init, fn)` | `(T, (T, Row) => T) => T` | Aggregate rows to single value |
| `.scan(init, fn)` | `(T, (T, Row) => T) => T[]` | Running aggregation, returns array |
| `.map(fn)` | `(Row => U) => U[]` | Transform each row |
| `.filter(fn)` | `(Row => bool) => Row[]` | Filter rows by predicate |
| `.find(fn)` | `(Row => bool) => Row?` | First row matching predicate |
| `.some(fn)` | `(Row => bool) => bool` | Any row matches |
| `.every(fn)` | `(Row => bool) => bool` | All rows match |

**Convenience aggregations:**
| Method | Description |
|--------|-------------|
| `.sum(fn)` | `reduce(0, (acc, r) => acc + fn(r))` |
| `.count(fn?)` | Count rows (optionally matching predicate) |
| `.max(fn)` | Maximum value of expression |
| `.min(fn)` | Minimum value of expression |
| `.avg(fn)` | Average value of expression |

**Example — sum with convenience method:**
```
tblSales.sum(r => r[Price] * r[Qty])
```
Equivalent to:
```
tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

---

### Transpilation Output

**Quick Mode input:**
```
tblSales.reduce(0, (acc, r) => acc + r[Price] * r[Qty])
```

**C# output (competition use):**
```csharp
public static object Udf_a1b2c3(object[,] tbl)
{
    var headers = BuildHeaderIndex(tbl);  // {"Price": 0, "Qty": 1, ...}
    double acc = 0;
    for (int i = 1; i < tbl.GetLength(0); i++)
    {
        var price = Convert.ToDouble(tbl[i, headers["Price"]]);
        var qty = Convert.ToDouble(tbl[i, headers["Qty"]]);
        acc = acc + price * qty;
    }
    return acc;
}
```

**Robust Mode input:**
```
=LET(
    price, tblSales[[#Headers],[Price]],
    qty, tblSales[[#Headers],[Qty]],
    tbl, tblSales[#All],
    `tbl.reduce(0, (acc, r) => acc + r.price * r.qty)`
)
```

**C# output (competition use):**
```csharp
public static object Udf_a1b2c3(object[,] tbl, string priceCol, string qtyCol)
{
    var headers = BuildHeaderIndex(tbl);
    double acc = 0;
    for (int i = 1; i < tbl.GetLength(0); i++)
    {
        var price = Convert.ToDouble(tbl[i, headers[priceCol]]);
        var qty = Convert.ToDouble(tbl[i, headers[qtyCol]]);
        acc = acc + price * qty;
    }
    return acc;
}
```

The LET-bound header references are passed as string parameters to the UDF.

---

### VBA Export

For portable submission, the same DSL transpiles to VBA:

```vba
Function Udf_a1b2c3(tbl As Range) As Variant
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
    
    Udf_a1b2c3 = acc
End Function
```

---

### Edge Cases

**Empty table:**
- `.reduce()` returns initial value
- `.scan()` returns empty array
- `.sum()`, `.avg()` return 0

**Missing column name:**
- Quick mode: Runtime error with clear message ("Column 'Pricee' not found. Available: Price, Qty, Region")
- Robust mode: Error at LET evaluation before DSL runs

**Mixed header types:**
- Numeric headers (e.g., years as column names) work with bracket syntax: `r[2024]`
- Empty header cells: accessible by index only

**Duplicate column names:**
- First occurrence wins for name-based access
- Use index-based access for subsequent columns
