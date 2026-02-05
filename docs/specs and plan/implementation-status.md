# Implementation Status

This document tracks implementation progress against the [Excel UDF Add-in Specification](excel-udf-addin-spec.md).

**Last updated:** 2025-02-05

---

## Quick Summary

| Area | Status | Notes |
|------|--------|-------|
| Core DSL (lexer, parser) | ✅ Complete | All token types, operators, lambdas |
| Cell operations (.cells, .values) | ✅ Complete | Object model + fast path |
| Row/column access (.rows, .cols) | ✅ Complete | Index-based access working |
| LINQ operations | ✅ Complete | where, select, orderBy, take, skip, distinct |
| Aggregations | ✅ Complete | sum, avg, min, max, count, first, last |
| .map() (shape-preserving) | ✅ Complete | |
| .groupBy() | ✅ Complete | With and without aggregator |
| .reduce() (values) | ✅ Complete | Named `.aggregate()` in code |
| Deep property access | ✅ Complete | c.Interior.ColorIndex, c.Font.Bold |
| Type system + validation | ✅ Complete | With typo suggestions |
| Null-safe access (?, ??) | ✅ Complete | |
| Escape hatch (@) | ✅ Complete | |
| Row-wise .reduce() with columns | ⏳ Not started | Core spec feature |
| .scan() (running reduction) | ⏳ Not started | |
| .find(), .some(), .every() | ⏳ Not started | |
| Named column access (r[Price]) | ⏳ Not started | |
| Table detection (ListObject) | ⏳ Not started | |
| .withHeaders() | ⏳ Not started | |
| LET robust column references | ⏳ Not started | |
| Statement lambdas | ⏳ Not started | |
| VBA transpiler | ⏳ Not started | Export feature |
| Floating editor | ⏳ Not started | Post-MVP |

**Legend:** ✅ Complete | 🚧 Partial | ⏳ Not started

---

## Detailed Status

### Lexer & Parser

| Feature | Lexer | Parser | Tests | Notes |
|---------|-------|--------|-------|-------|
| Identifiers | ✅ | ✅ | ✅ | |
| Numbers | ✅ | ✅ | ✅ | |
| Strings (with escapes) | ✅ | ✅ | ✅ | \n, \t, \r, \\, \" |
| Range references | ✅ | ✅ | ✅ | A1:B10, $A$1:$B$10 |
| Arithmetic operators | ✅ | ✅ | ✅ | +, -, *, / |
| Comparison operators | ✅ | ✅ | ✅ | ==, !=, >, <, >=, <= |
| Logical operators | ✅ | ✅ | ✅ | &&, \|\|, ! |
| Null coalescing (??) | ✅ | ✅ | ✅ | |
| Safe access suffix (?) | ✅ | ✅ | ✅ | obj.prop? |
| Escape hatch (@) | ✅ | ✅ | ✅ | obj.@prop |
| Lambda (=>) | ✅ | ✅ | ✅ | |
| Single-param lambda | ✅ | ✅ | ✅ | x => expr |
| Multi-param lambda | ✅ | ✅ | ✅ | (a, b) => expr |
| Statement lambda | ⏳ | ⏳ | ⏳ | x => { ... } |
| Method chains | ✅ | ✅ | ✅ | .where().select() |
| Index access | ✅ | ✅ | ✅ | arr[0] |
| Member access | ✅ | ✅ | ✅ | obj.prop |
| Object literals | ⏳ | ⏳ | ⏳ | {key: value} |

---

### Cell/Range Access

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| `.cells` | ✅ | ✅ | Object model path |
| `.values` | ✅ | ✅ | Fast value-only path |
| `.rows` | ✅ | ✅ | Returns object[][] |
| `.cols` | ✅ | ✅ | Returns object[][] |
| Implicit `.values` | ✅ | ✅ | Default when no .cells |
| Implicit `.toArray()` | ✅ | ✅ | Auto-materialise collections |

---

### LINQ-Style Operations

| Method | Transpiler | Tests | Notes |
|--------|------------|-------|-------|
| `.where(predicate)` | ✅ | ✅ | Filter |
| `.select(transform)` | ✅ | ✅ | Map to 1D |
| `.map(transform)` | ✅ | ✅ | Preserve 2D shape |
| `.orderBy(key)` | ✅ | ✅ | Sort ascending |
| `.orderByDesc(key)` | ✅ | ✅ | Sort descending |
| `.take(n)` | ✅ | ✅ | Supports negative n |
| `.skip(n)` | ✅ | ✅ | Supports negative n |
| `.distinct()` | ✅ | ✅ | Remove duplicates |
| `.groupBy(key)` | ✅ | ✅ | Group and flatten |
| `.groupBy(key, agg)` | ✅ | ✅ | Returns [key, value] pairs |
| `.reduce(seed, fn)` | ✅ | ✅ | Named `.aggregate()` in code |
| `.reduce(fn)` | ✅ | ✅ | First element as seed |
| `.toArray()` | ✅ | ✅ | Explicit materialisation |

---

### Aggregations

| Method | Transpiler | Tests | Notes |
|--------|------------|-------|-------|
| `.sum()` | ✅ | ✅ | |
| `.sum(selector)` | ✅ | ✅ | |
| `.avg()` / `.average()` | ✅ | ✅ | |
| `.min()` | ✅ | ✅ | With optional selector |
| `.max()` | ✅ | ✅ | With optional selector |
| `.count()` | ✅ | ✅ | |
| `.first()` | ✅ | ✅ | |
| `.firstOrDefault()` | ✅ | ✅ | |
| `.last()` | ✅ | ✅ | |
| `.lastOrDefault()` | ✅ | ✅ | |

---

### Row-Wise Table Operations

**This is the main area requiring implementation.**

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| `.reduce(init, fn)` on rows | ⏳ | ⏳ | Row object with column access |
| `.scan(init, fn)` | ⏳ | ⏳ | Running reduction |
| `.find(predicate)` | ⏳ | ⏳ | First matching row |
| `.some(predicate)` | ⏳ | ⏳ | Any row matches |
| `.every(predicate)` | ⏳ | ⏳ | All rows match |
| Row index access `r[0]` | ✅ | ✅ | Works in .rows.where() |
| Row named access `r[Price]` | ⏳ | ⏳ | Requires header detection |
| Row dot notation `r.Price` | ⏳ | ⏳ | Requires header detection |
| Negative index `r[-1]` | ⏳ | ⏳ | Last column |

---

### Table Detection & Headers

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| Excel Table detection | ⏳ | ⏳ | ListObject lookup |
| `.withHeaders()` modifier | ⏳ | ⏳ | First row as headers |
| Header index building | ⏳ | ⏳ | Column name → index map |
| Dynamic column lookup | ⏳ | ⏳ | Runtime, not hardcoded |

---

### LET Integration

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| Basic LET variable tracking | ✅ | ✅ | ExpressionContext |
| UDF naming from LET var | ✅ | ✅ | |
| Source preservation (_src_) | ⏳ | ⏳ | |
| Table binding detection | ⏳ | ⏳ | `tbl, tblSales` |
| Column binding detection | ⏳ | ⏳ | `price, tblSales[Price]` |
| Robust column param gen | ⏳ | ⏳ | Pass column names to UDF |
| Edit mode reconstruction | ⏳ | ⏳ | Ctrl+Shift+` |

---

### Cell Properties (Object Model)

| Property | Transpiler | Tests | Notes |
|----------|------------|-------|-------|
| `.value` | ✅ | ✅ | |
| `.color` | ✅ | ✅ | Interior.ColorIndex |
| `.rgb` | ✅ | ✅ | Interior.Color |
| `.bold` | ✅ | ✅ | Font.Bold |
| `.italic` | ✅ | ✅ | Font.Italic |
| `.fontSize` | ✅ | ✅ | Font.Size |
| `.format` | ✅ | ✅ | NumberFormat |
| `.formula` | ✅ | ✅ | |
| `.row` | ✅ | ✅ | |
| `.col` | ✅ | ✅ | |
| `.address` | ✅ | ✅ | |

---

### Deep Property Access

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| `c.Interior.ColorIndex` | ✅ | ✅ | |
| `c.Interior.Color` | ✅ | ✅ | |
| `c.Interior.Pattern` | ✅ | ✅ | |
| `c.Font.Bold` | ✅ | ✅ | |
| `c.Font.Italic` | ✅ | ✅ | |
| `c.Font.Size` | ✅ | ✅ | |
| `c.Font.Color` | ✅ | ✅ | |
| `c.Font.Name` | ✅ | ✅ | |
| Type validation | ✅ | ✅ | |
| Typo suggestions | ✅ | ✅ | Levenshtein distance |

---

### Null-Safe Access

| Feature | Transpiler | Tests | Notes |
|---------|------------|-------|-------|
| `obj.prop?` suffix | ✅ | ✅ | Try-catch wrapper |
| `??` operator | ✅ | ✅ | Null coalescing |
| `@` escape hatch | ✅ | ✅ | Bypass type validation |
| Combined `obj.@prop?` | ✅ | ✅ | |

---

### Statement Lambdas

| Feature | Lexer | Parser | Transpiler | Tests |
|---------|-------|--------|------------|-------|
| Detect `{` after `=>` | ⏳ | ⏳ | ⏳ | ⏳ |
| Brace-balanced capture | ⏳ | ⏳ | ⏳ | ⏳ |
| Emit as literal C# | ⏳ | ⏳ | ⏳ | ⏳ |

---

### Export & Portability

| Feature | Status | Notes |
|---------|--------|-------|
| C# transpiler | ✅ | Primary backend |
| VBA transpiler | ⏳ | Export feature |
| "Prepare for Export" | ⏳ | |
| Bake to values | ⏳ | |
| VBA injection | ⏳ | Requires trust settings |

---

### UI Features

| Feature | Status | Notes |
|---------|--------|-------|
| Quote-prefix detection | ✅ | SheetChange handler |
| Formula rewriting | ✅ | |
| Error display (#UDF_ERR) | ✅ | |
| Cell comment errors | ✅ | |
| Floating editor | ⏳ | Post-MVP |
| Ctrl+Shift+E shortcut | ⏳ | |
| Ctrl+Shift+N (name UDF) | ⏳ | |
| Autocomplete | ⏳ | Post-MVP |

---

### Built-in Algorithms (Post-MVP)

| Algorithm | Status | Notes |
|-----------|--------|-------|
| `.shortestPath()` | ⏳ | Dijkstra |
| `.connectedComponents()` | ⏳ | |
| `.topoSort()` | ⏳ | |
| `.permutations()` | ⏳ | |
| `.combinations(k)` | ⏳ | |
| `.subsets()` | ⏳ | |
| `.iterate(fn, until)` | ⏳ | |

---

## Implementation Priority

Based on spec and competitive Excel use cases:

### High Priority (Core Row Operations)
1. **Row-wise `.reduce()` with column access** — enables `tbl.reduce(0, (acc, r) => acc + r[Price] * r[Qty])`
2. **Named column access** (`r[Price]`, `r.Price`) — requires header detection
3. **Table detection** — recognise Excel Tables by name
4. **`.withHeaders()`** — enable named access for plain ranges

### Medium Priority (Enhanced Row Operations)
5. `.scan()` — running totals, state accumulation
6. `.find()`, `.some()`, `.every()` — row predicates
7. Negative index support (`r[-1]`)
8. LET robust column references

### Lower Priority (Polish & Export)
9. Statement lambdas
10. VBA transpiler
11. Source preservation pattern
12. Edit mode reconstruction

---

## Test Coverage Notes

**Well-tested areas:**
- Parser: 50+ tests
- Transpiler: 80+ tests
- Integration: ValuePathTests, ObjectModelTests

**Gaps:**
- No tests for row-wise reduce/scan/find/some/every (not implemented)
- No tests for named column access (not implemented)
- No tests for table detection (not implemented)
- No tests for statement lambdas (not implemented)
