# 0013 — Type System Cleanup

## Problem

The wrapper type system has gaps and inconsistencies that surface when users try to combine element-wise iteration with cell formatting access or primitive return types. The motivating use case:

```
maze.Map(c => c.Cell.Color)
```

This should return a 2D array of color values matching the shape of `maze`. Currently it requires:

```
maze.Map(c => new ExcelScalar(((ExcelScalar)c).Cell.Color))
```

...which still returns `#VALUE!` because Map doesn't thread cell access context through to its elements.

## Root Causes

### 1. Element-wise lambda parameters typed as `ExcelValue` instead of `ExcelScalar`

All element-wise operations (`Map`, `Where`, `Any`, `All`, `First`, `OrderBy`, `ForEach`, etc.) use `Func<ExcelValue, ...>`. But element-wise iteration always produces single-cell values, internally wrapped as `ExcelScalar`. The `ExcelValue` typing is a lie — users never receive an `ExcelArray` or `ExcelTable` as a lambda parameter in these operations.

This matters because `.Cell` (singular) is defined only on `ExcelScalar`. Users must cast to access cell formatting: `((ExcelScalar)c).Cell.Color`.

### 2. No implicit conversion into `ExcelScalar`

Implicit conversions **out** exist (`ExcelValue` → `double`, `string`, `bool`), but not **back** (`double`/`int`/`string`/`bool` → `ExcelScalar`). This forces users to write `new ExcelScalar(value)` instead of just returning the value from Map lambdas.

### 3. `Map()` creates origin-less ExcelScalars

`ExcelArray.Map()` wraps each element as `new ExcelScalar(_data[r, c])` — no `RangeOrigin`, no `CellAccessor`. So `.Cell` always throws `InvalidOperationException`, which surfaces as `#VALUE!`. Map has the position info (`r`, `c`) but doesn't pass it through. Compare this to how `Row.MakeScalar()` creates ExcelScalars with `CellAccessor` lambdas that capture position.

### 4. `Map()` has no generic overload for non-ExcelValue return types

`Map` requires `Func<ExcelValue, ExcelValue>` — returning a primitive won't compile. `Select` appears to work with primitives (e.g., `.Select(v => 1)`), but only because LINQ's `Enumerable.Select<TSource, TResult>` silently shadows the custom implementation via `IExcelRange : IEnumerable<ExcelValue>`. Map has no LINQ equivalent, so it needs its own generic overload.

### 5. Custom element-wise `Select` and `SelectMany` are redundant with LINQ

The custom `Select(Func<ExcelValue, ExcelValue>)` on `IExcelRange`/`ExcelValue`/`ExcelArray`/`ExcelScalar` is shadowed by LINQ whenever the return type isn't `ExcelValue`. Both paths produce the same single-column result. The custom version returns `IExcelRange` which theoretically enables chaining into `IExcelRange`-specific methods, but after Select the 2D shape is lost and those methods don't apply meaningfully. `SelectMany` has the same redundancy.

## Proposed Solution

### A. Change all element-wise lambda input parameters from `ExcelValue` to `ExcelScalar`

Every element-wise operation iterates single cells. The lambda parameter should reflect this:

| Operation | Current signature | New signature |
|---|---|---|
| `Map` | `Func<ExcelValue, ExcelValue>` | `Func<ExcelScalar, ExcelScalar>` |
| `Where` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `Any` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `All` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `First` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `FirstOrDefault` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `OrderBy` | `Func<ExcelValue, object>` | `Func<ExcelScalar, object>` |
| `OrderByDescending` | `Func<ExcelValue, object>` | `Func<ExcelScalar, object>` |
| `Aggregate` | `Func<ExcelValue, ExcelValue, ExcelValue>` | `Func<ExcelValue, ExcelScalar, ExcelValue>` |
| `Scan` | `Func<ExcelValue, ExcelValue, ExcelValue>` | `Func<ExcelValue, ExcelScalar, ExcelValue>` |
| `ForEach` | `Action<ExcelValue>` | `Action<ExcelScalar>` |
| `ForEach` (indexed) | `Action<ExcelValue, int, int>` | `Action<ExcelScalar, int, int>` |

This gives `.Cell` access in all element-wise lambdas without casting. The `IEnumerable<ExcelValue>` enumerator (used by LINQ) is unaffected — it still yields `ExcelValue`.

Note: for `Aggregate` and `Scan`, only the "current element" parameter changes to `ExcelScalar`. The accumulator stays `ExcelValue` since it can be any computed value.

### B. Add implicit conversions into `ExcelScalar`

Add implicit operators: `double` → `ExcelScalar`, `int` → `ExcelScalar`, `string` → `ExcelScalar`, `bool` → `ExcelScalar`.

This enables `Map(c => c.Cell.Color)` where `.Color` returns `int` — the int auto-wraps to `ExcelScalar` to satisfy the return type. It also makes `Map(c => (double)c * 2)` and similar patterns work naturally.

### C. Thread origin context through `Map()`

`ExcelArray.Map()` should create ExcelScalars with cell access context, following the same pattern as `Row.MakeScalar()`:
- When the source ExcelArray has a `RangeOrigin`, create a `CellAccessor` for each element that captures `(origin.SheetName, origin.TopRow + r, origin.LeftCol + c)`
- When no origin exists (computed arrays), create ExcelScalars without cell access (current behaviour)

This means `.Cell` works inside Map lambdas for original ranges but correctly throws for derived/computed arrays.

### D. Add generic `Map<TResult>` overload

Add `IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector)` alongside the existing `Map`. The generic version boxes each `TResult` directly into the result array, bypassing the need for `ExcelScalar` wrapping on the return side.

This mirrors what LINQ's `Select<TSource, TResult>` provides, but preserves Map's 2D-shape semantics.

### E. Remove custom element-wise `Select` and `SelectMany`

Remove the custom `Select(Func<ExcelValue, ExcelValue>)` and `SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>>)` from:
- `IExcelRange` interface
- `ExcelValue` abstract base
- `ExcelScalar` implementation
- `ExcelArray` implementation

LINQ provides these via `IEnumerable<ExcelValue>`. `ResultConverter` already handles `IEnumerable<T>` results.

**Not affected:** `RowCollection.Select`, `ColumnCollection.Select`, `GroupedRowCollection.Select` — these are on separate types with different signatures (`Func<dynamic, object>`) and are not part of this change.

## User Stories

- As a user, I want to write `maze.Map(c => c.Cell.Color)` and get a 2D array of color values, without casting or wrapping.
- As a user, I want to return primitive values from Map lambdas (e.g., `data.Map(c => (double)c * 2)`) without wrapping in `new ExcelScalar(...)`.
- As a user, I want `.Cell` accessible in Where/Any/All predicates (e.g., `data.Where(c => c.Cell.Bold)`) without casting.
- As a user, I want `.Select(v => someExpression)` to work regardless of the expression's return type, just as it does today via LINQ.

## Acceptance Criteria

- [ ] `maze.Map(c => c.Cell.Color)` compiles and returns a 2D array of ColorIndex values
- [ ] `data.Map(c => (double)c * 2)` compiles and returns a 2D array of doubled values
- [ ] `data.Where(c => c.Cell.Bold)` compiles — `.Cell` accessible without casting
- [ ] Implicit conversions into ExcelScalar work: `ExcelScalar s = 42;` compiles
- [ ] Custom element-wise Select removed — `data.Select(v => 1)` still works via LINQ fallback
- [ ] Custom element-wise SelectMany removed — LINQ fallback works
- [ ] Existing tests pass (with updates to reflect new signatures)
- [ ] `RowCollection.Select`, `ColumnCollection.Select`, `GroupedRowCollection.Select` are unaffected
- [ ] User spec (0005) updated to reflect ExcelScalar lambda parameters and Map generic overload

## Out of Scope

- Adding `.Cell` to `ExcelValue` base class — not needed since element-wise lambdas now receive `ExcelScalar` directly
- Adding origin threading to other transformation operators (Where, OrderBy, etc.) — these correctly drop origin because positions no longer map to original cells after filtering/reordering
- Changes to RowCollection/ColumnCollection/GroupedRowCollection Select methods

## Open Questions

None — resolved through discussion.
