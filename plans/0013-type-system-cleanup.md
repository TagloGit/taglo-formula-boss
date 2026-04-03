# 0013 — Type System Cleanup — Implementation Plan

## Overview

Fix the wrapper type system so element-wise lambdas correctly type their parameters as `ExcelScalar`, add implicit conversions into `ExcelScalar`, thread origin context through `Map()`, add a generic `Map<TResult>` overload, and remove the redundant custom `Select`/`SelectMany`.

Reference: [specs/0013-type-system-cleanup.md](../specs/0013-type-system-cleanup.md)

## Files to Touch

### Runtime library (formula-boss.Runtime)

| File | Changes |
|------|---------|
| `IExcelRange.cs` | Change lambda param types to `ExcelScalar`; remove `Select`/`SelectMany`; add `Map<TResult>` |
| `ExcelValue.cs` | Mirror interface changes on abstract declarations; remove `Select`/`SelectMany`; change `ForEach` param type |
| `ExcelScalar.cs` | Add implicit operators (double, int, string, bool → ExcelScalar); remove `Select`/`SelectMany` impl; update `Map` signature |
| `ExcelArray.cs` | Update all element-wise methods to pass `ExcelScalar`; remove `Select`/`SelectMany` impl; thread origin through `Map`; add `Map<TResult>` impl |
| `Row.cs` | Override `GetEnumerator` already yields `ExcelScalar` via `MakeScalar` — no change needed |

### Tests (formula-boss.Runtime.Tests)

| File | Changes |
|------|---------|
| `ExcelArrayTests.cs` | Update lambda params from `ExcelValue` casts to `ExcelScalar`; update Select tests to use LINQ; add Map origin/Cell test; add Map<TResult> test |
| `ExcelScalarTests.cs` | Update Select/SelectMany tests to use LINQ; update Map test; add implicit conversion tests |
| `ColumnTests.cs` | Update Select test to use LINQ |
| `ExcelTableTests.cs` | Update Rows.Select test (RowCollection.Select is unaffected, but verify) |
| `ColumnCollectionTests.cs` | Verify Select tests still work (ColumnCollection.Select is unaffected) |

### Integration tests

| File | Changes |
|------|---------|
| `formula-boss.IntegrationTests/WrapperTypePipelineTests.cs` | Verify all formula tests pass; update if any call custom Select directly |
| `formula-boss.AddinTests/PipelineTests.cs` | Add test for `maze.Map(c => c.Cell.Color)` |

### Spec

| File | Changes |
|------|---------|
| `specs/0005-formula-boss-user-spec.md` | Update element-wise table: lambda param is `ExcelScalar` not `ExcelValue`; add Map<TResult> |

## Order of Operations

### Step 1 — Implicit conversions into ExcelScalar

**Why first:** These are additive (no breaking changes) and are needed by Map return types in later steps.

- Add to `ExcelScalar.cs`:
  ```csharp
  public static implicit operator ExcelScalar(double value) => new(value);
  public static implicit operator ExcelScalar(int value) => new((double)value);
  public static implicit operator ExcelScalar(string value) => new(value);
  public static implicit operator ExcelScalar(bool value) => new(value);
  ```
- Add unit tests for each conversion in `ExcelScalarTests.cs`

### Step 2 — Change element-wise lambda input params to `ExcelScalar`

**Why:** Core type-correctness fix. This is a breaking change to the API signatures, so do it in one pass.

Changes across `IExcelRange.cs`, `ExcelValue.cs`, `ExcelScalar.cs`, `ExcelArray.cs`:

| Method | Old param | New param |
|--------|-----------|-----------|
| `Where` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `Any` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `All` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `First` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `FirstOrDefault` | `Func<ExcelValue, bool>` | `Func<ExcelScalar, bool>` |
| `Map` | `Func<ExcelValue, ExcelValue>` | `Func<ExcelScalar, ExcelScalar>` |
| `OrderBy` | `Func<ExcelValue, object>` | `Func<ExcelScalar, object>` |
| `OrderByDescending` | `Func<ExcelValue, object>` | `Func<ExcelScalar, object>` |
| `ForEach` (simple) | `Action<ExcelValue>` | `Action<ExcelScalar>` |
| `ForEach` (indexed) | `Action<ExcelValue, int, int>` | `Action<ExcelScalar, int, int>` |

**Not changed:** `Aggregate` and `Scan` — these already use `Func<dynamic, dynamic, dynamic>`, so `.Cell` is already accessible through dynamic dispatch.

**Internal changes:**
- `ExcelArray.ElementWise()` already returns `IEnumerable<ExcelScalar>` — no change needed
- `ExcelArray.GetEnumerator()` calls `ElementWise()` — returns `ExcelScalar` instances typed as `ExcelValue` (covariant). No change needed.
- `ExcelValue.ForEach(Action<ExcelValue>)` is a concrete method using `foreach (var el in this)` — change param to `Action<ExcelScalar>` and cast: `action((ExcelScalar)el)` (safe because GetEnumerator always yields ExcelScalar).

Update all unit tests that cast lambda params (e.g., `(double)v` → `(double)c`, `new ExcelScalar((double)v * 10)` → `c * 10` or `new ExcelScalar((double)c * 10)`).

### Step 3 — Thread origin context through Map

**Why:** Enables `.Cell` access inside Map lambdas.

In `ExcelArray.Map()`, change from:
```csharp
var mapped = selector(new ExcelScalar(_data[r, c]));
```

To (following the `Row.MakeScalar` / `ExcelArray.Rows` pattern):
```csharp
Func<Cell>? cellAccessor = _origin != null && RuntimeBridge.GetCell != null
    ? () => RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow + r, _origin.LeftCol + c)
    : null;
var scalar = new ExcelScalar(_data[r, c]) { CellAccessor = cellAccessor };
var mapped = selector(scalar);
```

**Important:** Must capture `r` and `c` in the closure correctly. Since `r` and `c` are loop variables, the lambda captures the variable, not the value. Need to use local copies:
```csharp
var localR = r;
var localC = c;
Func<Cell>? cellAccessor = _origin != null && RuntimeBridge.GetCell != null
    ? () => RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow + localR, _origin.LeftCol + localC)
    : null;
```

Add unit test in `ExcelArrayTests.cs` that verifies `.Cell` is accessible inside a Map lambda when the source array has a `RangeOrigin`.

### Step 4 — Add generic Map\<TResult\> overload

**Why:** Enables `Map(c => c.Cell.Color)` where Color returns int, without needing implicit conversion to ExcelScalar on the return side.

Add to `IExcelRange.cs`:
```csharp
IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector);
```

Add to `ExcelValue.cs` (abstract):
```csharp
public abstract IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector);
```

Implement in `ExcelArray.cs`:
```csharp
public override IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector)
{
    var rows = _data.GetLength(0);
    var cols = _data.GetLength(1);
    var result = new object?[rows, cols];
    for (var r = 0; r < rows; r++)
        for (var c = 0; c < cols; c++)
        {
            var localR = r;
            var localC = c;
            Func<Cell>? cellAccessor = _origin != null && RuntimeBridge.GetCell != null
                ? () => RuntimeBridge.GetCell(_origin.SheetName, _origin.TopRow + localR, _origin.LeftCol + localC)
                : null;
            var scalar = new ExcelScalar(_data[r, c]) { CellAccessor = cellAccessor };
            result[r, c] = selector(scalar);
        }
    return new ExcelArray(result, ColumnMap);
}
```

Implement in `ExcelScalar.cs`:
```csharp
public override IExcelRange Map<TResult>(Func<ExcelScalar, TResult> selector)
{
    var result = selector(this);
    return result is ExcelValue ev ? (IExcelRange)ev : new ExcelScalar(result);
}
```

Add tests: `Map_Generic_ReturnsInt`, `Map_Generic_ReturnsString`, `Map_Generic_Preserves2DShape`.

**Note:** The non-generic `Map(Func<ExcelScalar, ExcelScalar>)` is still needed because when the return type IS ExcelScalar, we want to unwrap `.RawValue` (existing behaviour) rather than boxing the ExcelScalar object itself. The generic version boxes `TResult` directly.

### Step 5 — Remove custom Select and SelectMany

**Why:** Redundant with LINQ. Simplifies the API.

Remove from:
- `IExcelRange.cs` — remove `Select` and `SelectMany` declarations
- `ExcelValue.cs` — remove abstract `Select` and `SelectMany`
- `ExcelScalar.cs` — remove `Select` and `SelectMany` implementations
- `ExcelArray.cs` — remove `Select` and `SelectMany` implementations

**Verify LINQ fallback works:** `data.Select(v => 1)` should resolve to `Enumerable.Select<ExcelValue, int>`, return `IEnumerable<int>`, and `ResultConverter` handles it via the generic `IEnumerable` path (lines 124-157 in ResultConverter.cs).

Update tests:
- Tests that called `.Select()` with `ExcelValue` return type will now hit LINQ. The result type changes from `IExcelRange` to `IEnumerable<ExcelValue>`. Tests that cast the result to `ExcelArray` or check `.RawValue` will need updating to work with `IEnumerable<T>` instead.
- `SelectMany` tests similarly need updating.

### Step 6 — Update user spec and run full test suite

- Update `specs/0005-formula-boss-user-spec.md`:
  - Element-wise operations table: "The lambda parameter is `ExcelScalar`" (was `ExcelValue`)
  - Add note about `Map<TResult>` generic overload
  - Add example: `maze.Map(c => c.Cell.Color)`
- Run `dotnet test formula-boss/formula-boss.slnx` — all tests must pass
- Run AddIn tests to verify end-to-end

## Testing Approach

### Unit tests (formula-boss.Runtime.Tests)

- **Implicit conversions:** `double` → `ExcelScalar`, `int` → `ExcelScalar`, `string` → `ExcelScalar`, `bool` → `ExcelScalar`
- **Map with origin:** Create ExcelArray with RangeOrigin + mock RuntimeBridge.GetCell, verify `.Cell` accessible in Map lambda
- **Map\<TResult\>:** Verify generic Map returns 2D array of raw values (int, string, etc.)
- **LINQ Select fallback:** Verify `data.Select(v => 1)` returns correct results through ResultConverter
- **Existing tests updated:** ~45 tests need lambda param updates (casts change from `(double)v` to `(double)c` etc.)

### Integration tests (formula-boss.IntegrationTests)

- Existing WrapperTypePipelineTests should pass with no formula changes (user formulas don't reference `ExcelValue` explicitly)

### AddIn tests (formula-boss.AddinTests)

- **New test:** Formula using `maze.Map(c => c.Cell.Color)` on a colored range — the motivating use case
- Existing tests should pass unchanged (user formulas use implicit typing)

## Risks

1. **LINQ Select resolution ambiguity:** After removing custom Select, if a user writes `.Select(v => v * 2)`, does LINQ resolve the return type as `ExcelValue` (via operator overload) or `double`? It should be `ExcelValue` since the `*` operator on ExcelValue returns ExcelValue. But worth verifying.

2. **Closure variable capture in Map:** The `r` and `c` loop variables must be captured as local copies in the CellAccessor lambda. Forgetting this would cause all cells to reference the last position.

3. **ExcelScalar.Map non-generic:** Currently returns `selector(this)` as `IExcelRange`. With the signature change to `Func<ExcelScalar, ExcelScalar>`, the return is `ExcelScalar` which implements `IExcelRange`, so this still works.
