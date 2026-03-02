# 0003 — Wrapper-Type Architecture — Implementation Plan

**Spec:** `specs/0003-wrapper-type-architecture.md`
**Epic:** #59
**Review notes:** `plans/0003-architecture-review-notes.md`

## Context

The architecture review produced 13 findings from a line-by-line walkthrough of generated UDF code, runtime types, and pipeline components. The spec has been updated to incorporate all design decisions. This plan describes surgical changes to the existing code on branch `62-transpiler-rewrite-v2`, grouped into PR-sized units in dependency order.

## Order of Operations

### PR 1: Shared ToResult and Fixed GetHeaders Contract

**Issue label:** `refactor`

Consolidates duplicate result conversion logic and fixes the GetHeaders contract mismatch (Findings 1, 3, 5).

**Files:**
- `formula-boss.Runtime/ResultConverter.cs` — add `public static object Convert(object? result)` with the full dispatch chain. Change all `ToResult` overloads to return `object` (bare scalars for single values, `object[,]` for multi-cell).
- `formula-boss/RuntimeHelpers.cs` — change `GetHeadersDelegate` type from `Func<object, string[]?>` to `Func<object[,], string[]?>` (accepts already-extracted values, not raw ExcelReference).
- `formula-boss/AddIn.cs` — `ToResultDelegate` body becomes `result => ResultConverter.Convert(result)`. Fix `GetHeadersDelegate` to accept `object[,]` (remove internal `GetValuesFromReference` call).
- `formula-boss/Transpilation/CodeEmitter.cs` — generated header extraction code passes `{input}__values` (cast to `object[,]`) to `GetHeadersDelegate`, not `{input}__raw`.
- `formula-boss.IntegrationTests/NewPipelineTestHelpers.cs` — replace `ToResultDelegate` and `GetHeadersDelegate` with calls to shared implementations.

**Verification:** All existing `WrapperTypePipelineTests` pass. Scalar results now return bare values (test assertions updated).

---

### PR 2: Runtime Types — IExcelRange on ExcelValue, Element-wise Methods, RowCollection

**Issue label:** `refactor`

Core type system changes (Findings 2, 6, 10, 11). This is the largest PR.

**Files:**
- `formula-boss.Runtime/IExcelRange.cs` — change all method signatures from `Func<Row, ...>` to `Func<ExcelValue, ...>`. Change `Rows` property type to `RowCollection`.
- `formula-boss.Runtime/ExcelValue.cs` — add `: IExcelRange` to class declaration. Declare all `IExcelRange` members as `abstract`.
- `formula-boss.Runtime/ExcelArray.cs` — reimplement all methods for element-wise iteration (cell-by-cell, row-major). `Rows` returns `new RowCollection(...)`. Aggregations iterate cells directly.
- `formula-boss.Runtime/ExcelScalar.cs` — reimplement with single-element semantics. `Rows` returns single-row `RowCollection`.
- `formula-boss.Runtime/RowCollection.cs` (**new**) — class with `List<Row>` backing. Instance methods with `Func<dynamic, ...>` parameters: `Where`, `Select`, `Any`, `All`, `First`, `FirstOrDefault`, `OrderBy`, `OrderByDescending`, `Count`, `Take`, `Skip`, `ToRange`.
- `formula-boss/AddIn.cs` — add `RuntimeBridge.GetCell` initialization (Finding 11).

**Verification:** Update test assertions for scalar results. Add tests for `RowCollection` operations and element-wise `Any`/`Where`. Test expressions using `.Rows` now go through `RowCollection`.

---

### PR 3: Pipeline Simplification — Flat Parameters, Free-Variable-Only Detection

**Issue label:** `refactor`

Removes explicit lambda syntax, primary input concept, old column binding mechanism, and three-category parameter model (Findings 6, 7, 8, 9, 12).

**Files:**
- `formula-boss/Transpilation/InputDetector.cs` — remove outer-lambda detection, `ExtractPrimaryInput`, `IsSugarSyntax`. All inputs via free variable analysis. `DetectionResult` changes: remove `IsSugarSyntax`, `Inputs`, `FreeVariables`, `HasStringBracketAccess`; add `Parameters: IReadOnlyList<string>` (flat ordered list) and `HeaderVariables: IReadOnlySet<string>` (per-variable tracking of which vars need header extraction).
- `formula-boss/Transpilation/CodeEmitter.cs` — remove `FindArrowIndex`. Uniform preamble for every parameter. Header extraction conditioned on `detection.HeaderVariables.Contains(param)`. Remove `(IExcelRange)` cast (ExcelValue now implements IExcelRange).
- `formula-boss/Interception/FormulaPipeline.cs` — `PipelineResult`: remove `InputParameter`, `ColumnParameters`, `AdditionalInputs`, `FreeVariables`; add `Parameters: IReadOnlyList<string>?`. Remove `_columnParamsCache`, `_additionalInputsCache`.
- `formula-boss/Interception/LetFormulaRewriter.cs` — delete `InjectHeaderBindings`, `ColumnParameter` record. `ProcessedBinding`: flatten to `Parameters`. `AppendUdfCall`: single `string.Join(", ", processed.Parameters)`.
- `formula-boss/Interception/FormulaInterceptor.cs` — update `ProcessLetFormula` to use flat `Parameters`. Remove `columnBindings` extraction. Update `ProcessBacktickFormula` UDF call construction. Simplify `ExpressionContext`.

**Verification:** Update test expressions (remove `(tbl) =>` prefixes). Add tests for per-variable header extraction. Verify LET formula wiring with flat parameter list.

---

### PR 4: Dot-Notation-to-Bracket Rewrite

**Issue label:** `enhancement`

Implements the dot notation intellisense and auto-rewrite (spec "RowCollection" section). Users type `r.Population2025`, get intellisense from a synthetic typed Row class, and the transpiler rewrites to `r["Population 2025"]` before compilation.

**Files:**
- `formula-boss/Transpilation/ColumnMapper.cs` (**new**) — builds sanitised→original mapping from column headers. `Sanitise(string columnName) → string` removes spaces/special chars. `BuildMapping(string[] headers) → Dictionary<string, string>`. Detects conflicts (two columns → same sanitised name) and excludes those.
- `formula-boss/Transpilation/DotNotationRewriter.cs` (**new**) — Roslyn syntax rewriter. Walks the expression AST, finds `MemberAccessExpressionSyntax` where the identifier matches a sanitised column name, rewrites to `ElementAccessExpressionSyntax` with the original column name as a string literal.
- `formula-boss/Transpilation/CodeEmitter.cs` — after detecting free variables and before emitting the method body, run `DotNotationRewriter` on the expression using the column mapping from headers of each `HeaderVariable`.
- Intellisense synthetic document builder (future, but `ColumnMapper` provides the foundation) — synthetic Row class with real properties generated from the mapping.

**Verification:** Test `r.Population2025` → `r["Population 2025"]` rewrite. Test conflict detection. Test mixed dot and bracket access in same expression.

---

### PR 5: Test Cleanup and Acceptance Tests

**Issue label:** `enhancement`

Final test updates and spec acceptance criteria coverage.

**Files:**
- `formula-boss.IntegrationTests/NewPipelineTestHelpers.cs` — final cleanup, ensure all helpers use shared implementations.
- `formula-boss.IntegrationTests/WrapperTypePipelineTests.cs` — add acceptance criteria tests:
  - Multi-input population filter with maxPop
  - Nested lambda: `pConts.Any(c => c == r["Continent"])` inside `.Rows.Where()`
  - Scalar return is bare value (not 1x1 array)
  - Dot notation rewrite with column containing spaces
  - Statement block with multiple free variables

**Verification:** Full test suite green. All spec acceptance criteria covered.

## Critical Files Summary

| File | PRs | Key Change |
|---|---|---|
| `formula-boss.Runtime/IExcelRange.cs` | 2 | `Func<Row,...>` → `Func<ExcelValue,...>` |
| `formula-boss.Runtime/ExcelValue.cs` | 2 | Add `: IExcelRange`, abstract members |
| `formula-boss.Runtime/ExcelArray.cs` | 2 | Element-wise iteration, RowCollection |
| `formula-boss.Runtime/ResultConverter.cs` | 1 | Return `object`, shared `Convert()` |
| `formula-boss.Runtime/RowCollection.cs` | 2 | **New** — `Func<dynamic,...>` instance methods |
| `formula-boss/Transpilation/InputDetector.cs` | 3 | Free-var-only, flat `Parameters`, per-var `HeaderVariables` |
| `formula-boss/Transpilation/CodeEmitter.cs` | 1,3,4 | Uniform preamble, dot notation rewrite |
| `formula-boss/Interception/FormulaPipeline.cs` | 3 | Flat `Parameters` on `PipelineResult` |
| `formula-boss/Interception/LetFormulaRewriter.cs` | 3 | Remove column binding, flat params |
| `formula-boss/Interception/FormulaInterceptor.cs` | 3 | Use flat `Parameters` |
| `formula-boss/AddIn.cs` | 1,2 | Shared ToResult, fixed GetHeaders, init GetCell |
| `formula-boss/Transpilation/ColumnMapper.cs` | 4 | **New** — sanitised↔original column name mapping |
| `formula-boss/Transpilation/DotNotationRewriter.cs` | 4 | **New** — Roslyn syntax rewriter for dot→bracket |
