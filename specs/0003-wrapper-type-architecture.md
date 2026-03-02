# 0003 — Wrapper-Type Architecture

## Problem

The current Formula Boss transpiler converts a custom DSL into generated C# code via string-based transformations. This creates several compounding problems:

1. **LET variable capture fails** — identifiers not in `_lambdaParameters` fall through to `__source__`, so external LET variables are inaccessible inside lambdas (#53, #54).
2. **Dynamic dispatch breaks nested lambdas** — captured LET variables arrive as `dynamic`, and C# cannot resolve lambda arguments against dynamic dispatch (#56).
3. **DSL vs C# boundary confuses users** — `.some()` works in DSL expressions but not in statement expressions (where `.Any()` is needed). Users don't know which methods work where (#57).
4. **Intellisense is limited** — the custom `CompletionProvider` only covers DSL methods. No completions for C# string methods, LINQ, regex inside statement expressions (#58).
5. **Single-input limitation** — UDFs accept one implicit `__source__` parameter. Multi-table operations require workarounds.

These issues share a root cause: the transpiler generates code by string manipulation rather than operating on real types.

## Proposed Solution

Replace DSL syntax transformation with a **pre-compiled type-safe wrapper library**. Formula Boss expressions become standard C# operating on typed facades (`ExcelTable`, `ExcelArray`, `ExcelScalar`) that wrap Excel's raw `object`/`object[,]` values.

The transpiler's role simplifies to:
1. Detect all external identifiers via free variable analysis
2. Emit a C# method with typed wrapper parameters
3. Pass the user's code through mostly as-is

The DSL vs C# boundary disappears. Users write C# everywhere. The magic is in the types they operate on, not in syntax transformation.

---

## User Stories

- As a power user, I want to filter a table using LET variables (scalars, arrays) inside my lambda so that I can write complex multi-criteria filters without workarounds.
- As a power user, I want consistent method names (`.Any()`, `.Where()`, `.All()`) regardless of where I use them so that I don't have to remember two sets of APIs.
- As a power user, I want intellisense for C# string methods, LINQ, and regex inside my expressions so that I can discover available methods without consulting documentation.
- As a power user, I want to operate on multiple inputs in a single expression so that I can combine data from different tables or ranges.

---

## Design

### Expression Model

A Formula Boss expression is a **C# expression or statement block that references named Excel values.** All external identifiers are detected automatically via free variable analysis — there is no explicit parameter syntax.

#### Method chain (the 95% case)

```
`tblCountries.Rows.Where(r => r["Population"] > 1000)`
```

The transpiler detects `tblCountries` as a free variable (the only unaccounted-for identifier) and makes it a UDF parameter.

#### Multiple inputs

```
`tblCountries.Rows.Where(r => r["Population"] < maxPop)`
```

Both `tblCountries` and `maxPop` are detected as free variables. Both become UDF parameters. There is no distinction between "primary" and "secondary" inputs — all parameters are equal.

#### Statement block

```
`{
    var count = tblCountries.Rows.Count();
    if (count > someThreshold)
        return tblCountries.Rows.Where(r => r["Price"] > 5).ToResult();
    return someOtherTable.Sum();
}`
```

`tblCountries`, `someThreshold`, and `someOtherTable` are all detected as free variables.

#### Automatic LET variable capture

In a LET formula, free variables are resolved by Excel in the LET scope:

```
=LET(
    maxPop, XLOOKUP(...),
    pConts, TEXTSPLIT(...),
    result, `tblCountries.Rows.Where(r =>
        r["Population"] < maxPop
        && pConts.Any(c => c == r["Continent"]))`,
    result)
```

`tblCountries`, `maxPop`, and `pConts` are all detected as free variables and become UDF parameters. The LET formula rewriter wires each LET variable to the corresponding UDF argument.

**No explicit lambda input syntax.** The `(input1, input2) => expression` form is not supported. All inputs are detected via free variable analysis. This eliminates the confusing requirement that parameter names must match LET variable names exactly.

### Type System

#### Wrapper Type Hierarchy

```
ExcelValue : IExcelRange                — base wrapper for any Excel value
  ├─ ExcelTable                         — named table (ListObject), has column metadata
  ├─ ExcelArray                         — raw object[,], always 2D (from TEXTSPLIT, SORT, ranges, etc.)
  └─ ExcelScalar                        — single value with single-element collection semantics
```

`ExcelValue` itself implements `IExcelRange`. This means every parameter — regardless of whether it wraps a table, array, or scalar — supports the full range API without casting. The concrete subclasses override the `IExcelRange` methods with type-appropriate behaviour:

- `ExcelScalar` has single-element semantics: `.Any()` tests the one value, `.Where()` returns 0 or 1 elements, `.Count()` returns 1, etc.
- `ExcelArray` iterates element-wise over the 2D array in row-major order.
- `ExcelTable` inherits from `ExcelArray`, adding column metadata.

Because `ExcelValue` implements `IExcelRange`, generated code never needs an `(IExcelRange)` cast. `ExcelValue.Wrap()` returns an `ExcelValue`, and all range methods are available directly.

#### IExcelRange Interface

Methods directly on `IExcelRange` iterate **element-wise** over all values (row-major: left-to-right, top-to-bottom). The lambda parameter is `ExcelValue`, not `Row`:

```csharp
public interface IExcelRange
{
    RowCollection Rows { get; }       // iterate as typed rows (see RowCollection below)
    ColumnCollection Cols { get; }    // iterate as columns
    CellCollection Cells { get; }     // iterate as COM cells (forces IsMacroType)

    // Element-wise operations (lambda receives each cell as ExcelValue)
    IExcelRange Where(Func<ExcelValue, bool> predicate);
    IExcelRange Select(Func<ExcelValue, ExcelValue> selector);
    IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector);
    bool Any(Func<ExcelValue, bool> predicate);
    bool All(Func<ExcelValue, bool> predicate);
    ExcelValue First(Func<ExcelValue, bool> predicate);
    ExcelValue FirstOrDefault(Func<ExcelValue, bool> predicate);

    // Aggregations
    int Count();
    ExcelScalar Sum();
    ExcelScalar Min();
    ExcelScalar Max();
    ExcelScalar Average();

    // Shape-preserving transform
    IExcelRange Map(Func<ExcelValue, ExcelValue> selector);

    // Sorting
    IExcelRange OrderBy(Func<ExcelValue, ExcelValue> keySelector);
    IExcelRange OrderByDescending(Func<ExcelValue, ExcelValue> keySelector);

    // Partitioning
    IExcelRange Take(int count);     // negative = TakeLast
    IExcelRange Skip(int count);     // negative = SkipLast
    IExcelRange Distinct();

    // Folding
    ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);
    IExcelRange Scan(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);

    // Grouping
    // GroupBy TBD — needs design for key/group result types
}
```

**Element-wise vs row-based iteration:** Methods directly on `IExcelRange` (`.Where()`, `.Any()`, `.Select()`, etc.) iterate over individual cell values as `ExcelValue`. To iterate row-by-row, use `.Rows` which returns a `RowCollection` with its own set of methods that take `Func<dynamic, ...>` (see RowCollection below).

Example:
- `pConts.Any(c => c == r["Continent"])` — `c` is an `ExcelValue`, iterates all cells
- `tbl.Where(v => v > 5)` — `v` is an `ExcelValue`, iterates all cells element-wise
- `tbl.Rows.Where(r => r["Price"] > 5)` — `r` is a `Row` (via `dynamic`), iterates row-by-row

#### RowCollection

`.Rows` returns a `RowCollection` — a custom collection type with **instance methods** that accept `Func<dynamic, ...>` parameters. This is necessary because:

1. LINQ extension methods cannot accept `Func<dynamic, ...>` (CS1977)
2. Instance methods can — the lambda parameter `r` is typed as `dynamic`
3. Dynamic dispatch on `Row` (which extends `DynamicObject`) triggers `TryGetMember`, but this only works at runtime, not for intellisense

**Column access uses dot notation with auto-rewrite:** The user types `r.Population2025` in the editor and gets intellisense completions from a synthetic typed Row class. Before compilation, the transpiler rewrites dot access to bracket access: `r.Population2025` → `r["Population 2025"]`. This gives the UX of dot notation (discoverability, no quotes/brackets to type) with the reliability of bracket access at runtime.

**How the dot-notation-to-bracket rewrite works:**

1. **Header extraction:** When table metadata is available (from LET context or structured references), the transpiler knows the column names — e.g. `["Country", "Population 2025", "GDP"]`.
2. **Sanitised property mapping:** Each column name is sanitised to a valid C# identifier by removing spaces and special characters — e.g. `"Population 2025"` → `Population2025`. A reverse mapping is stored: `{ "Population2025": "Population 2025" }`.
3. **Synthetic Row class for intellisense:** The Roslyn synthetic document includes a typed Row subclass with real properties:
   ```csharp
   class _Row_ {
       public ColumnValue Country => this["Country"];
       public ColumnValue Population2025 => this["Population 2025"];
       public ColumnValue GDP => this["GDP"];
   }
   ```
   This gives full intellisense — the user types `r.P` and sees `Population2025` as a completion.
4. **Pre-compilation rewrite:** Before Roslyn compiles the expression, a rewrite pass uses the mapping to convert all `r.SanitisedName` accesses to `r["Original Name"]`. The compiled code only uses the `Row.this[string]` indexer.
5. **Conflict detection:** If two columns sanitise to the same identifier (e.g. `Foo Bar` and `FooBar` both → `FooBar`), the ambiguous name is excluded from the synthetic class. The user must use bracket access for those columns. This should be rare.

**Bracket access also works directly:** Users can always write `r["Population 2025"]` or `r[0]` without going through dot notation. Bracket access bypasses the rewrite step entirely.

```csharp
public class RowCollection
{
    // Instance methods — Func<dynamic, ...> enables r.Col and r["Col"] syntax
    public RowCollection Where(Func<dynamic, bool> predicate);
    public IExcelRange Select(Func<dynamic, ExcelValue> selector);
    public bool Any(Func<dynamic, bool> predicate);
    public bool All(Func<dynamic, bool> predicate);
    public Row First(Func<dynamic, bool> predicate);
    public Row? FirstOrDefault(Func<dynamic, bool> predicate);
    public RowCollection OrderBy(Func<dynamic, object> keySelector);
    public RowCollection OrderByDescending(Func<dynamic, object> keySelector);

    // Non-lambda methods
    public int Count();
    public RowCollection Take(int count);
    public RowCollection Skip(int count);
    public RowCollection Distinct();

    // Conversion
    public IExcelRange ToRange();  // convert back to ExcelArray for further element-wise ops
}
```

#### Row Type

When iterating via `.Rows`, each element is a `Row`:

```csharp
public class Row : DynamicObject
{
    // Bracket access (always works, used at runtime after dot-notation rewrite)
    public ColumnValue this[string columnName] { get; }
    public ColumnValue this[int index] { get; }     // zero-based, negative for last

    // Dynamic member access (DynamicObject.TryGetMember)
    // Enables r.Population at runtime; intellisense provided via synthetic typed Row class
}
```

#### ColumnValue Type

Each column access returns a `ColumnValue` that supports both value operations and cell escalation:

```csharp
public class ColumnValue
{
    // Implicit conversion to common types for natural comparisons
    public static implicit operator double(ColumnValue v);
    public static implicit operator string(ColumnValue v);
    public static implicit operator bool(ColumnValue v);

    // Comparison operators
    public static bool operator >(ColumnValue a, ColumnValue b);
    public static bool operator <(ColumnValue a, ColumnValue b);
    // ... ==, !=, >=, <=

    // Cross-type operators (ColumnValue vs ExcelValue)
    public static bool operator >(ColumnValue a, ExcelValue b);
    public static bool operator <(ColumnValue a, ExcelValue b);
    // ... enables r["Price"] > maxPop where maxPop is an ExcelScalar

    // Arithmetic operators
    public static ColumnValue operator +(ColumnValue a, ColumnValue b);
    public static ColumnValue operator -(ColumnValue a, ColumnValue b);
    public static ColumnValue operator *(ColumnValue a, ColumnValue b);
    public static ColumnValue operator /(ColumnValue a, ColumnValue b);

    // Cell escalation — access formatting properties
    public Cell Cell { get; }

    // Raw value access
    public object Value { get; }
}
```

The user writes `r["Population"] > 1000` and it works via operator overloading. If they write `r["Population"].Cell.Color`, the static analyzer detects `.Cell` usage and sets `IsMacroType = true` on the UDF registration.

#### ExcelScalar Type

Wraps a single value with smart operator overloading:

```csharp
public class ExcelScalar : ExcelValue
{
    // Same implicit conversions and operators as ColumnValue
    // Enables: someLetVariable > 250 (regardless of underlying type)

    // String methods available when wrapping a string
    // (via implicit conversion to string, then standard C# string methods work)
}
```

#### ExcelArray Type

Wraps `object[,]` from TEXTSPLIT, SORT, spill ranges, etc. **Always 2D** — even a 1×N horizontal or N×1 vertical array is stored as `object[,]`.

```csharp
public class ExcelArray : ExcelValue
{
    // All IExcelRange methods (element-wise)
    // Element-wise iteration yields ExcelScalar elements
    // .Rows returns RowCollection for row-by-row iteration
    // .Cols iterates column-by-column
    // For 1×N or N×1 arrays, all iteration modes produce equivalent results

    // Enables: pConts.Any(c => c == r["Continent"])
    // where pConts is an ExcelArray from TEXTSPLIT

    // .Select() flattens to 1D (maps elements to values)
    // .Map() preserves 2D shape
}
```

#### Cell Type (COM access)

```csharp
public class Cell
{
    public int Color { get; }         // Interior.ColorIndex
    public int Rgb { get; }           // Interior.Color
    public bool Bold { get; }         // Font.Bold
    public bool Italic { get; }       // Font.Italic
    public double FontSize { get; }   // Font.Size
    public string Format { get; }     // NumberFormat
    public string Formula { get; }    // Cell formula
    public int Row { get; }           // Row number
    public int Col { get; }           // Column number
    public string Address { get; }    // Cell address
    public object Value { get; }      // Cell value

    // Sub-objects for deep property access
    public Interior Interior { get; }
    public Font Font { get; }
}
```

### Comparison Operators

Comparison operators on `ExcelValue` currently return `bool`. This is correct for the common case (`maxPop > 1000`, `r["Price"] > threshold`).

**Future: element-wise array comparison.** In Excel, `A1:A5 > 5` returns a spilled array of TRUE/FALSE. To support `tbl > 5` returning an array of booleans, operators would need to return `ExcelValue` (not `bool`) and delegate to virtual instance methods:

```csharp
public static ExcelValue operator >(ExcelValue a, double b) => a.CompareGreaterThan(b);
// ExcelScalar overrides → returns ExcelScalar(bool)
// ExcelArray overrides → returns ExcelArray of booleans
```

Plus `operator true`/`operator false` on `ExcelValue` so scalar results still work in `if` conditions. Array results in boolean context should throw.

**This is deferrable.** Changing return type from `bool` to `ExcelValue` later is backwards compatible as long as implicit conversion to `bool` and `operator true`/`false` exist. No blocking dependency on current work. Implement with `bool`-returning operators for now.

### Value/Cell Access Model

Access to cell formatting (color, bold, etc.) requires COM interop, which is slower and requires `IsMacroType = true`. The wrapper types support **lazy cell escalation**:

- **Default path (fast):** Operations on values. `r["Population"] > 1000` works via `ColumnValue` operators, no COM needed.
- **Cell path (when needed):** Access `.Cell` on any `ColumnValue` or use `.Cells` on a range. `r["Population"].Cell.Color == 6` escalates to COM.
- **Static detection:** Before compilation, the transpiler scans for `.Cell` / `.Cells` usage to determine whether `IsMacroType = true` is needed.

Shorthand accessors on `IExcelRange`:
- `.Rows` — iterate as `Row` objects via `RowCollection` (value access + `.Cell` available)
- `.Cols` — iterate as column arrays
- `.Cells` — iterate as `Cell` objects directly (always COM, forces `IsMacroType`)
- Direct methods (no accessor) — element-wise on values

### Method Naming

All methods use **C# LINQ naming conventions**:

| Old DSL | New (C# convention) |
|---|---|
| `.some()` | `.Any()` |
| `.every()` | `.All()` |
| `.find()` | `.FirstOrDefault()` |
| `.where()` | `.Where()` |
| `.select()` | `.Select()` (row-major element iteration, flattens to 1D) |
| (new) | `.SelectMany()` (standard LINQ — flatten nested collections) |
| `.map()` | `.Map()` (Formula Boss-specific, preserves 2D shape) |
| `.reduce()` / `.aggregate()` | `.Aggregate()` |
| `.scan()` | `.Scan()` |
| `.orderBy()` | `.OrderBy()` |
| `.orderByDesc()` | `.OrderByDescending()` |
| `.take()` | `.Take()` |
| `.skip()` | `.Skip()` |
| `.distinct` | `.Distinct()` |
| `.count` | `.Count()` |
| `.sum()` / `.avg` / `.min` / `.max` | `.Sum()` / `.Average()` / `.Min()` / `.Max()` |

Formula Boss-specific names that have no C# equivalent:
- `.Rows`, `.Cols`, `.Cells` — accessor properties on `IExcelRange`
- `.Map()` — shape-preserving transform (distinct from `.Select()` which flattens)
- `.Scan()` — running fold (not in standard LINQ)
- `.WithHeaders()` — mark first row as headers

### Transpiler Changes

The transpiler simplifies significantly. Its job becomes:

1. **Detect free variables** — parse the expression with Roslyn, identify all unaccounted-for identifiers (not inner lambda params, not C# keywords, not type names, not method names, not locally declared variables). These become UDF parameters in a single flat ordered list.
2. **Detect cell usage** — scan for `.Cell` / `.Cells` to set `IsMacroType`
3. **Emit C# code:**
   ```csharp
   public static object UdfName(object tblCountries__raw, object maxPop__raw)
   {
       // Every parameter gets the same preamble:
       // 1. Check for ExcelReference
       // 2. Extract values if reference
       // 3. Extract headers if this variable uses string bracket access
       // 4. Wrap with ExcelValue.Wrap()
       var tblCountries__isRef = tblCountries__raw?.GetType()?.Name == "ExcelReference";
       var tblCountries__values = tblCountries__isRef == true
           ? FormulaBoss.RuntimeHelpers.GetValuesFromReference(tblCountries__raw)
           : tblCountries__raw;
       string[]? tblCountries__headers = /* header extraction if tblCountries uses r["Col"] */;
       var tblCountries = ExcelValue.Wrap(tblCountries__values, tblCountries__headers);

       var maxPop = ExcelValue.Wrap(maxPop__raw);  // no string bracket access → no headers

       // User's code, mostly verbatim
       var __result = tblCountries.Rows.Where(r => r["Population"] < maxPop);
       return __result.ToResult();
   }
   ```
4. **Result conversion** — `.ToResult()` returns `object`: bare scalars for single values, `object[,]` for multi-cell results. The UDF return type is `object`, which ExcelDNA handles correctly for both cases.

**Uniform parameter treatment:** Every parameter (regardless of how it was detected) gets the same wrapping preamble. The only per-variable distinction is whether to extract headers — determined by whether **that specific variable** is used with string bracket access in the expression (e.g. `r["Col"]` where `r` comes from that variable's `.Rows`), not a global boolean.

**No "primary input" concept.** All parameters are equal. `table1.Sum() + table2.Sum()` and `table1 + table2` both work — the transpiler doesn't need to find a "primary" input.

### Variable Naming

Formula Boss does not impose any naming restrictions on variables beyond what Excel and C# already require:
- **External inputs** (LET variables, named ranges, cell references) follow Excel's naming rules. Formula Boss just detects and passes them through.
- **Local variables** in statement blocks (`var i = 0`) follow C# naming rules. These are excluded from free variable detection.
- **Inner lambda parameters** (`r` in `.Where(r => ...)`) follow C# naming rules. Excluded from free variable detection.

Single-cell references like `A1` or `B2` are valid C# identifiers, so they pass through as free variables. Excel resolves them as cell references. This works correctly but by design rather than explicit support.

### Intellisense: Roslyn Completion Service

Replace the custom `CompletionProvider` with Roslyn's `CompletionService`:

1. **Build a synthetic C# document** from the user's code:
   ```csharp
   using System;
   using System.Linq;
   using System.Text.RegularExpressions;
   using FormulaBoss.Runtime;  // pre-compiled wrapper types

   class _Synthetic_ {
       void _M_(ExcelValue tblCountries, ExcelValue pConts, ExcelValue maxPop) {
           /* user's code here, caret position mapped */
       }
   }
   ```
   All parameters are typed as `ExcelValue` (which implements `IExcelRange`), so all range methods are available. When table metadata is known, the synthetic document includes a typed Row subclass with real properties for dot-notation intellisense (see RowCollection section).

2. **Map the caret position** from the editor to the synthetic document
3. **Call `CompletionService.GetCompletionsAsync()`** at the mapped position
4. **Filter and present results** — prioritize Formula Boss types, add descriptions

This provides:
- Wrapper type method completions (`.Rows`, `.Where()`, `.Any()`, etc.)
- String methods (`.Split()`, `.Contains()`, `.Trim()`, etc.)
- LINQ methods on any `IEnumerable<>`
- Regex completions
- Smart parameter info
- Full context awareness (knows the type at each chain position)
- Column name completions via dot notation on synthetic typed Row class

**Performance:** Roslyn completion is designed for interactive use in Visual Studio — fast enough for keystroke-by-keystroke completion. The synthetic document is small, so compilation overhead is minimal.

**Variable type inference for the synthetic document:**
- Table-bound LET variables → `ExcelTable` (with column metadata if known)
- Array-producing LET variables (TEXTSPLIT, SORT, etc.) → `ExcelArray`
- Scalar LET variables (XLOOKUP returning single value, literal) → `ExcelScalar`
- Unknown → `ExcelValue` (base type, all methods available)

### Pre-compiled Wrapper Assembly

The wrapper types are **compiled as part of the add-in**, not generated per-UDF. This is critical for:

- **Performance** — no per-UDF compilation cost for the type library
- **Intellisense** — Roslyn can reference the types as known assemblies
- **Testing** — wrapper types are unit-testable independently of Excel

**Assembly identity constraint:** The wrapper types must NOT reference ExcelDNA types. They work with `object`, `object[,]`, and COM interop via `dynamic`/reflection. This avoids the assembly identity mismatch documented in CLAUDE.md.

The wrapper assembly is referenced by generated UDF code via Roslyn's `MetadataReference`. The `ExcelValue.Wrap()` factory inspects runtime values and returns the appropriate typed wrapper.

### Breaking Changes

This is a pre-release product. Existing formulas using old DSL syntax (`.some()`, `.where()` lowercase, etc.) will break. No migration path is provided — users rewrite expressions using the new C# method names.

---

## Acceptance Criteria

- [ ] Wrapper types (`ExcelTable`, `ExcelArray`, `ExcelScalar`, `Row`, `ColumnValue`, `Cell`) are pre-compiled and unit-tested
- [ ] `ExcelValue` implements `IExcelRange` — no casting needed in generated code
- [ ] Free variable detection works for all inputs: `tblCountries.Rows.Where(r => r["Pop"] > maxPop)` detects both `tblCountries` and `maxPop`
- [ ] LET variable capture works automatically for scalars, arrays, and tables
- [ ] Nested lambdas resolve correctly: `pConts.Any(c => c == r["Continent"])` where `pConts` is an `ExcelArray`
- [ ] Cell escalation works: `r["Population"].Cell.Color` triggers `IsMacroType` and returns formatting data
- [ ] Method naming is C# convention everywhere (`.Any()`, `.Where()`, `.All()`, etc.)
- [ ] `IExcelRange` methods iterate element-wise with `Func<ExcelValue, ...>`; `.Rows` returns `RowCollection` with `Func<dynamic, ...>`
- [ ] `ToResult()` returns bare scalars for single values, `object[,]` for multi-cell results
- [ ] Dot notation with intellisense: `r.P` offers `Population2025`, rewritten to `r["Population 2025"]` before compilation
- [ ] Roslyn-powered intellisense provides completions for wrapper types AND standard C# methods
- [ ] The target formula from #59 compiles and returns correct results
- [ ] Source preservation and error reporting continue to work

## Out of Scope

- VBA transpiler updates (Phase 12, #21) — will need separate work to understand wrapper types
- Built-in algorithm library (#22) — future work, but wrapper types provide the foundation
- Case-tolerant method names — rely on intellisense to prevent casing errors
- GroupBy detailed design — noted as TBD in the type system, will be designed during implementation
- Element-wise array comparison operators — deferrable, approach documented in Comparison Operators section
- Old column binding mechanism (`ColumnParameter`, header injection) — replaced by dot notation with auto-rewrite and bracket access

## Design Decisions (Resolved)

1. **ExcelValue.Wrap() factory** — **Runtime type detection.** `Wrap()` inspects the runtime value and returns the appropriate type: `ExcelTable` for ListObject references, `ExcelArray` for `object[,]`, `ExcelScalar` for scalars. This is the safest approach as it handles Excel's inconsistencies (e.g. a formula sometimes returning a scalar, sometimes an array). `ExcelScalar` implements `IExcelRange` with single-element semantics so that calling `.Any()` on a scalar works naturally rather than throwing.

2. **Row column access** — **Dot notation with auto-rewrite.** Users type `r.Population2025` and get intellisense from a synthetic typed Row class built from table headers. Before compilation, the transpiler rewrites dot access to bracket access (`r.Population2025` → `r["Population 2025"]`) using a sanitised-name-to-original mapping. Bracket access (`r["Col"]`, `r[0]`) also works directly. Conflicts (two columns sanitising to the same identifier) are detected and those columns fall back to bracket-only access.

3. **Roslyn workspace management** — **Persistent `AdhocWorkspace`** created at add-in startup, reused across all completion requests. The workspace holds one project with wrapper type references. Each completion request updates only the synthetic document. Memory footprint is small (~few MB). Disposed on shutdown.

4. **ExcelArray shape** — **Always 2D.** `ExcelArray` wraps `object[,]` and preserves its shape. All arrays are conceptually 2D. For a 1×N or N×1 array, element-wise iteration (`.Where()` directly), `.Rows` iteration, and `.Cols` iteration produce equivalent results. `.Select()` flattens to 1D (mapping elements to values). `.Map()` preserves 2D shape.

5. **ExcelValue implements IExcelRange** — Eliminates the need for `(IExcelRange)` casts in generated code. Every `ExcelValue` supports the full range API. Concrete subclasses override with type-appropriate behaviour.

6. **All parameters are equal** — No "primary input" concept. All external identifiers detected via free variable analysis become UDF parameters in a single flat ordered list. The transpiler treats every parameter identically: check for ExcelReference, extract values, optionally extract headers (per-variable, based on whether that variable's rows use string bracket access), wrap with `ExcelValue.Wrap()`.

7. **Result conversion returns `object`** — `ToResult()` returns bare scalars for single values, `object[,]` for multi-cell results. The UDF return type is `object`. ExcelDNA handles both correctly — bare scalars display as single-cell values, `object[,]` spills as arrays.

8. **Header extraction is per-variable** — The transpiler tracks which specific variables are used with string bracket access (e.g. `tbl.Rows.Where(r => r["Col"])`) and only extracts headers for those variables. This replaces the previous global `HasStringBracketAccess` boolean.
