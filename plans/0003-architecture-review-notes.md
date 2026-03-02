# 0003 — Architecture Review Notes

## Instructions for Next Planner

This document contains 13 findings from a line-by-line walkthrough of the generated UDF code, Runtime types, and pipeline components against `specs/0003-wrapper-type-architecture.md`. Several findings represent design decisions that override or clarify the current spec.

**Your task has three phases:**

### Phase 1: Update the spec

Read every finding below and compare with `specs/0003-wrapper-type-architecture.md`. Where the findings contain design decisions, update the spec to match. Key decisions:

- Remove explicit lambda input syntax `(x, y) => ...` — all inputs detected via free variable analysis (Finding 12)
- Remove the "primary input" concept — all parameters are equal (Finding 12)
- `IExcelRange` methods iterate element-wise with `Func<ExcelValue, ...>`, NOT `Func<Row, ...>` (Finding 10)
- `.Rows` returns a collection with its own row-based methods (Finding 10)
- `ExcelValue` should implement `IExcelRange` directly — no cast needed (hierarchy discussion)
- Old column binding mechanism (`ColumnParameter`, header injection) is dead code — remove from spec (Finding 7)
- No Formula Boss-specific variable naming restrictions (Finding 13)
- Element-wise array comparison operators are deferrable but the approach is documented (Finding 4)

### Phase 2: Review current code against updated spec

The current code is on branch `issue-62-transpiler-rewrite`. Review each component and identify what needs to change:

- **Runtime types** (`formula-boss.Runtime/`): Structurally sound but need targeted fixes. `ExcelValue` needs `IExcelRange`, element-wise iteration needs adding, `ToResult` needs fixing for scalars.
- **Pipeline** (`formula-boss/Transpilation/`, `formula-boss/Interception/`): Needs more significant rework. CodeEmitter, InputDetector, and LetFormulaRewriter all have assumptions that the findings invalidate.
- **Delegates** (`AddIn.cs`, `RuntimeHelpers.cs`): Duplicate implementations need consolidating. `RuntimeBridge.GetCell` needs initialising.
- **Tests** (`formula-boss.IntegrationTests/`): Test helper has its own delegate implementations that drift from production.

### Phase 3: Plan the changes

Plan as surgery on the existing branch, not a rewrite. The Runtime types, ALC loading, test infrastructure, and COM wiring all work and should be preserved. Group changes by dependency order — some findings are prerequisites for others (e.g. `ExcelValue` implementing `IExcelRange` simplifies CodeEmitter changes).

Create GitHub Issues for each logical unit of work under epic #59.

---

## Findings

Findings from line-by-line walkthrough of generated code vs spec vs runtime types.

## Finding 1: Duplicate ToResultDelegate implementations

**Location:** `AddIn.cs:140` and `NewPipelineTestHelpers.cs:118`

Two independent implementations of the same result conversion logic. Both are big if/else chains handling `ExcelValue`, `IExcelRange`, `IEnumerable<Row>`, `IEnumerable<ColumnValue>`, primitives, etc.

**Risk:** They drift apart. A bug fix in one doesn't propagate to the other. Tests pass but Excel breaks (or vice versa).

**Root cause:** The delegate bridge pattern was needed before the ALC fix because generated code couldn't call Runtime methods directly. Now that generated code loads into ExcelDNA's ALC, it *can* call `ResultConverter.ToResult()` directly.

**Immediate fix:** Extract the delegate body into a common static method (in Runtime or a shared location). Both `AddIn.cs` and test setup call that method. One implementation, no drift.

**Longer-term question:** The ALC spike allows generated code to reference Runtime types directly (e.g. `ExcelValue.Wrap()`, `IExcelRange`). If ALC loading proves fully reliable, delegates like `ToResultDelegate` could potentially be replaced by direct calls. However, ALC reliability is not yet fully verified — some scenarios still return #VALUE!. Do NOT eliminate delegates until ALC is proven across all scenarios.

## Finding 2: Dot notation for column names (`r.Population`) doesn't compile in LINQ lambdas

**Location:** `Row.cs` extends `DynamicObject`, but LINQ binds lambda params statically.

`tbl.Rows.Where(r => r.Population > 1000)` fails to compile because `.Rows` returns `IEnumerable<Row>`, LINQ's `.Where()` binds `r` as the static type `Row`, and there's no `Population` property — only `DynamicObject.TryGetMember()` which requires `dynamic` dispatch.

**Why the old version worked:** The old transpiler did string transformation, rewriting `r.Population` into `r[columnIndex]` before Roslyn saw it. Dot notation was never compiled as C#.

**Key insight:** Instance methods on a custom collection type can accept `Func<dynamic, ...>` parameters. Extension methods cannot (CS1977). So if `.Rows` returns a `RowCollection` with instance methods like `.Where(Func<dynamic, bool>)`, then `r` is typed as `dynamic` and `r.Population` triggers `TryGetMember` at runtime.

**Option A — Bracket access only:** Users write `r["Population"]` or `r[0]`. Works today (once header runtime bug is fixed). No dot notation.

**Option B — Custom `RowCollection` with `Func<dynamic, ...>` instance methods:** `.Rows` returns `RowCollection` instead of `IEnumerable<Row>`. Methods like `.Where()`, `.Select()`, `.OrderBy()` are instance methods accepting `Func<dynamic, ...>`. Internally they delegate to LINQ on the underlying rows. Dot notation works.

**Option C — Generate typed Row classes per UDF:** Code generator emits a Row subclass with real properties per table. Roslyn sees real properties, dot notation and intellisense both work. More code generation complexity.

**Untested question:** Can a `static IEnumerable<dynamic> Where(this IEnumerable<dynamic> source, Func<dynamic, bool> predicate)` extension method work, or does CS1977 still apply? Worth a quick spike but the safe path is instance methods (Option B).

## Finding 3: `ToResult()` wraps scalars in 1x1 arrays unnecessarily

**Location:** `ResultConverter.cs` — all overloads return `object?[,]`

`ToResult(this int value)` returns `new object?[,] { { value } }`. Same for `double`, `string`, `bool`, `ExcelScalar`. This means scalar results like `tbl.Rows.Count()` get returned to Excel as a 1x1 array instead of a bare `3`.

The UDF return type is `object`, which can hold both scalars and arrays. ExcelDNA handles both correctly — bare scalars display as single-cell values, `object[,]` spills as arrays.

**Fix:** The common result conversion should return `object` (not `object?[,]`). Scalars return bare values. Only genuine multi-cell results return `object[,]`. This also means `ResultConverter.ToResult()` overloads for primitives can be removed or changed to return `object`.

## Finding 4: Element-wise comparison operators on arrays (future design, deferrable)

**Context:** In Excel, `A1:A5 > 5` returns a spilled array of TRUE/FALSE. Should `tbl > 5` do the same?

**Problem:** C# operators are static, not virtual — they resolve at compile time based on declared type. If the variable is typed as `ExcelValue`, the compiler always picks `ExcelValue.operator>` regardless of whether the runtime value is scalar or array.

**Solution pattern:** Operators on `ExcelValue` delegate to virtual instance methods:
```csharp
public static ExcelValue operator >(ExcelValue a, double b) => a.CompareGreaterThan(b);
// ExcelScalar overrides → returns ExcelScalar(bool)
// ExcelArray overrides → returns ExcelArray of booleans
```

Plus `operator true`/`operator false` on `ExcelValue` so scalar results still work in `if` conditions. Array results in boolean context should throw ("Cannot use array comparison as boolean condition").

**Is it deferrable?** Yes. Changing return type from `bool` to `ExcelValue` later is backwards compatible as long as implicit conversion to `bool` and `operator true`/`false` exist. No blocking dependency on current work.

**Decision:** Implement with `bool`-returning operators for now. Revisit element-wise array comparison later.

## Finding 5: GetHeadersDelegate has different contracts in AddIn vs tests

**Location:** `AddIn.cs:74` and `NewPipelineTestHelpers.cs:210`

The two implementations expect different input types:
- **Test version:** Expects `object[,]` directly. Casts and reads first row.
- **AddIn.cs version:** Expects an `ExcelReference`. Calls `GetValuesFromReference(rangeRef)` internally to extract the `object[,]`, then reads first row.

Same problem as Finding 1 — two implementations that can drift. But also a contract mismatch: the generated code calls `GetHeadersDelegate` in both the reference and non-reference branches, passing `tblSpaces__raw` either way. The AddIn.cs delegate will throw in the non-reference branch because `GetValuesFromReference` rejects non-`ExcelReference` inputs.

Additionally, in the reference branch, `GetValuesFromReference` is called **twice** — once in the generated code to populate `tblSpaces__values`, and again inside `GetHeadersDelegate` to extract headers. Redundant and potentially fragile.

**This is likely related to the `r["Price"]` #VALUE! bug.** The delegate silently returns null on exception (catch block at line 95), so if anything goes wrong in header extraction, headers are null, `Wrap` creates an `ExcelArray` instead of `ExcelTable`, and string key access fails at runtime. The error is swallowed.

**Proposed fix:** Same as Finding 1 — extract to a common implementation. The delegate should accept `object[,]` (already-extracted values), not a raw reference. The generated code already has the values — pass them to the header delegate instead of making it re-extract.

## Finding 6: CodeEmitter wrongly assumes syntactic role determines runtime type

**Location:** `CodeEmitter.cs:117-128`

**Root cause:** The code conflates how a variable was detected (syntactic role) with what Excel will pass at runtime (type). It has three tiers:

1. **First input** (identifier before first `.`) → full reference handling, header extraction, cast to `IExcelRange`
2. **Additional inputs** (other explicit lambda params) → reference handling but no cast, stays as `ExcelValue`
3. **Free variables** (LET variables used in body) → bare `ExcelValue.Wrap()`, no reference handling at all

But at runtime, any parameter could be a table, an array, a scalar, or an `ExcelReference`. The transpiler can't know — it depends on what the LET formula binds the variable to. For example:
- A free variable `pConts` from TEXTSPLIT is an array, not a scalar
- A free variable `maxNum` bound to a cell is an `ExcelReference`, not a value
- A second explicit input could be a table needing `.Rows`

**Fix:** Every parameter (input or free variable) gets the same preamble: check for `ExcelReference`, extract values, let `ExcelValue.Wrap()` determine the right wrapper at runtime. The only per-variable distinction is whether to extract headers — and that should be based on whether **that specific variable** uses string key access in the expression, not a global boolean.

**Related:** If `ExcelValue` implements `IExcelRange` directly (see hierarchy discussion), the cast to `IExcelRange` becomes unnecessary for all inputs, further simplifying this.

## Finding 7: Old column binding mechanism is dead code — remove it

**Location:** `LetFormulaRewriter.cs` (InjectHeaderBindings, ColumnParameter), `FormulaPipeline.cs` (ColumnParameters, UsedColumnBindings), `ProcessedBinding.ColumnParameters`

The old transpiler had a mechanism where LET variables bound to table columns (e.g. `p = tblSales[Price]`) were detected, and the rewriter injected `_price_hdr, INDEX(tblSales[[#Headers],[Price]],1)` into the LET formula. This allowed the old DSL to rewrite `r.price` to use the header value, which survived column renames.

The new wrapper type architecture doesn't use this. Column access is via `r["Price"]` or `r[0]` compiled directly into the generated code. The old `ColumnParameter`, `InjectHeaderBindings`, `ColumnBindingInfo`, and related plumbing should be removed.

**Column rename survivability** is handled by standard LET patterns if needed:
```
=LET(p, tblSales[[#Headers],[Price]],
  result, `tbl.Rows.Select(r => r[p])`,
  result)
```
Here `p` is a free variable passed as an `ExcelScalar` wrapping the string `"Price"`. If the column is renamed, the structured reference updates automatically.

**Note:** `Row` currently only has `this[string]` and `this[int]` indexers. For `r[p]` to work where `p` is an `ExcelScalar`, either an `ExcelValue` indexer overload is needed, or the user casts: `r[(string)p]`.

## Finding 8: Simplify parameter model — no need for three categories

**Location:** `PipelineResult` (InputParameter, AdditionalInputs, FreeVariables), `ProcessedBinding`, `AppendUdfCall`, `CodeEmitter`

The pipeline currently separates UDF parameters into three categories:
- `InputParameter` — single string, the "primary" input
- `AdditionalInputs` — list, other explicit lambda params
- `FreeVariables` — list, LET variables used in body

All three end up as `object` parameters in the UDF signature and arguments in the LET UDF call. The only requirement is that the order matches between the two. The three-way split adds complexity for no benefit.

**Simplify to one flat ordered list** of all parameters. CodeEmitter and LetFormulaRewriter just need an ordered list of names to emit. See Finding 12 — with explicit lambda syntax removed, there's no distinction between "inputs" and "free variables" from the code generation perspective. Everything is a detected identifier that becomes a UDF parameter.

## Finding 9: Free variable detection is "everything else" — developer should be aware of edge cases

The InputDetector doesn't identify what free variables *are* — it rules out everything it recognises (inputs, lambda params, keywords, type names, method names, local declarations, `__` prefixes) and whatever's left becomes a free variable. The LET rewriter emits the name as a UDF argument, and Excel resolves it in the LET scope.

This means free variables can be LET variables, named ranges, or even cell references (like `B2`) — anything Excel can resolve. The system doesn't validate them.

**Edge cases to be aware of:**
- Single-cell refs like `A1` or `B2` are valid C# identifiers, so they pass through as free variables. Excel resolves them as cell references. This works by accident, not by design.
- No Formula Boss-specific naming restrictions — see Finding 13.
- `HasStringBracketAccess` is global — doesn't track which variable uses string bracket access. All inputs get header extraction if any one uses it (see Finding 6).

## Finding 12: Remove explicit lambda input syntax — use free variable detection only

**Decision:** Users never need to write `(input1, input2) => expression`. All inputs are detected automatically via free variable analysis.

**How it works:**
- Sugar syntax: `table1.Rows.Where(r => r[0] > threshold)` — `table1` is primary input (before first `.`), `threshold` is a free variable. Both become UDF parameters.
- Statement blocks: `{ return table1.Rows.Count() + table2.Rows.Count(); }` — both `table1` and `table2` are free variables.
- Inner lambda params (`r` in `.Where(r => ...)`) are already excluded by `CollectAllLambdaParameters`.

**What this simplifies:**
- Remove `IsSugarSyntax` / explicit lambda distinction in InputDetector
- Remove `FindArrowIndex` in CodeEmitter
- Remove `AdditionalInputs` from pipeline (see Finding 8)
- All parameters get the same wrapping treatment (see Finding 6)
- Non-LET path in FormulaInterceptor no longer needs special multi-input handling

**Edge case:** If a user writes `(table1, table2) => ...` the InputDetector currently treats `table1` and `table2` as explicit inputs. With this change, the `(table1, table2) =>` prefix would need to either be a parse error or be silently stripped. Recommend: don't support it at all — it's confusing because the names must match LET variables exactly, which breaks normal lambda mental model.

**Primary input concept should be removed.** Currently `ExtractPrimaryInput` walks the leftmost member-access chain to find the "primary" input. Expressions without a member-access chain (e.g. `table1 + table2`) throw "Could not detect input identifier". With free-variable-only detection, there's no need for a primary input — all unaccounted-for identifiers become parameters equally. `table1.Sum() + table2.Sum()` and `table1 + table2` both work.

## Finding 13: No Formula Boss-specific variable naming restrictions

Formula Boss does not impose any naming restrictions on variables beyond what Excel and C# already require:
- **External inputs** (LET variables, named ranges, cell references) follow Excel's naming rules. These are pre-existing items — Formula Boss just detects and passes them through.
- **Local variables** in statement blocks (`var i1 = 0`) follow C# naming rules. These are excluded from free variable detection via `VariableDeclaratorSyntax` collection.
- **Inner lambda parameters** (`r` in `.Where(r => ...)`) follow C# naming rules. Excluded via `CollectAllLambdaParameters`.

No additional restrictions are needed.

## Finding 10: IExcelRange methods should iterate element-wise, not by Row

**Location:** `IExcelRange.cs`, `ExcelArray.cs`, `ExcelScalar.cs`

The spec says methods directly on `IExcelRange` (`Where`, `Any`, `Select`, etc.) iterate **element-wise** over all values (row-major: left-to-right, top-to-bottom), with lambda parameter typed as `ExcelValue`. Row-based iteration is only available via `.Rows`.

The current implementation has these methods taking `Func<Row, bool>` — meaning there's no way to iterate over individual values without going through `.Rows` and indexing.

**Spec intent:**
- `pConts.Any(c => c == r.Continent)` — `c` is an `ExcelValue`, iterates all cells
- `tbl.Where(v => v > 5)` — `v` is an `ExcelValue`, iterates all cells element-wise
- `tbl.Rows.Where(r => r["Price"] > 5)` — `r` is a `Row`, iterates row-by-row

**Fix:** `IExcelRange` methods should take `Func<ExcelValue, ...>` for element-wise iteration. `.Rows` returns a collection with its own `Where`, `Any`, etc. that take `Func<Row, ...>` (or `Func<dynamic, ...>` if we go with the RowCollection approach from Finding 2).

## Finding 11: RuntimeBridge.GetCell is never initialised — cell escalation is non-functional

**Location:** `RuntimeBridge.cs` declares `GetCell`, but `AddIn.AutoOpen` never sets it.

The entire cell escalation chain is wired up structurally:
- `ExcelArray.Rows` creates cell resolvers if `_origin` and `GetCell` exist
- `Row.MakeColumnValue` passes `CellAccessor` through
- `ColumnValue.Cell` invokes `CellAccessor`
- `ExcelArray.Cells` checks `GetCell` before yielding

But `GetCell` is always null, so every `.Cell` or `.Cells` access throws `InvalidOperationException`.

**Fix:** `AddIn.AutoOpen` needs to initialise `RuntimeBridge.GetCell` with a lambda that uses COM interop to read cell properties. This is a delegate bridge (safe pattern) — the delegate signature uses only primitives and `Cell` (a Runtime type with no COM dependency). Note that `RuntimeBridge.GetCell` uses `Cell` in its return type — this works because `Cell` is in `FormulaBoss.Runtime` which is loaded in the same ALC context.
