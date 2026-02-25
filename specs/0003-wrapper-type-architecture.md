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

Replace DSL syntax transformation with a **pre-compiled type-safe wrapper library**. Formula Boss expressions become standard C# lambdas operating on typed facades (`ExcelTable`, `ExcelArray`, `ExcelScalar`) that wrap Excel's raw `object`/`object[,]` values.

The transpiler's role simplifies to:
1. Identify inputs (from the expression and/or LET context)
2. Emit a C# lambda with typed wrapper parameters
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

A Formula Boss expression is a **C# lambda that takes named Excel values as inputs and returns a result.**

#### Single-input (sugar syntax — the 95% case)

```
`tblCountries.Rows.Where(r => r.Population > 1000)`
```

The transpiler infers a single input (`tblCountries`) and wraps it. Equivalent to:

```
`(tblCountries) => tblCountries.Rows.Where(r => r.Population > 1000)`
```

#### Multi-input (explicit lambda)

```
`(tblCountries, maxPop) => tblCountries.Rows.Where(r => r.Population < maxPop)`
```

Each identifier in the lambda parameter list becomes a UDF parameter, wrapped in the appropriate typed facade at runtime.

#### Multi-input with statement block

```
`(tblCountries, someOtherTable) => {
    if (tblCountries.Rows.Count() > 0)
        return "Countries!";
    return "Nothing";
}`
```

#### Automatic LET variable capture

In a LET formula, non-lambda identifiers used inside the expression are automatically detected and added as UDF parameters:

```
=LET(
    maxPop, XLOOKUP(...),
    pConts, TEXTSPLIT(...),
    result, `tblCountries.Rows.Where(r =>
        r.Population < maxPop
        && pConts.Any(c => c == r.Continent))`,
    result)
```

Here `tblCountries` is the primary input, and `maxPop` and `pConts` are automatically captured — no explicit multi-input syntax needed. The multi-input syntax is only required when the expression isn't a method chain on a single primary input.

### Type System

#### Wrapper Type Hierarchy

```
ExcelValue                            — base wrapper for any Excel value
  ├─ ExcelTable : IExcelRange         — named table (ListObject), has column metadata
  ├─ ExcelArray : IExcelRange         — raw object[,], always 2D (from TEXTSPLIT, SORT, ranges, etc.)
  └─ ExcelScalar : IExcelRange        — single value with single-element collection semantics
```

`ExcelScalar` implements `IExcelRange` with single-element semantics: `.Any()` tests the one value, `.Where()` returns 0 or 1 elements, `.Count()` returns 1, etc. This means users never hit a type boundary when a formula sometimes returns a scalar and sometimes an array.

#### IExcelRange Interface

Shared by `ExcelTable` and `ExcelArray`:

```csharp
public interface IExcelRange
{
    RowCollection Rows { get; }       // iterate as typed rows
    ColumnCollection Cols { get; }    // iterate as columns
    CellCollection Cells { get; }     // iterate as COM cells (forces IsMacroType)

    // Element-wise operations on values (row-major iteration: left-to-right, top-to-bottom)
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

**Note:** The exact generic signatures above are illustrative. The actual implementation will need to balance type safety with the `dynamic`/`object` nature of Excel data. For example, `Where` on `Rows` takes `Func<Row, bool>`, not `Func<ExcelValue, bool>`. The implementation will use appropriate generic type parameters or overloads per collection type.

#### Row Type

When iterating via `.Rows`, each element is a `Row` with named and indexed column access:

```csharp
public class Row
{
    // Named access (single-word columns)
    public ColumnValue Population => this["Population"];

    // Bracket access
    public ColumnValue this[string columnName] { get; }
    public ColumnValue this[int index] { get; }     // zero-based, negative for last

    // Dynamic member access for dot notation (r.Population)
    // Implemented via DynamicObject or source generation
}
```

Named column properties are either:
- Generated at compile time (if column names are known from table metadata or LET bindings)
- Resolved dynamically via `DynamicObject` fallback

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

The user writes `r.Population > 1000` and it works via operator overloading. If they write `r.Population.Cell.Color`, the static analyzer detects `.Cell` usage and sets `IsMacroType = true` on the UDF registration.

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
public class ExcelArray : ExcelValue, IExcelRange
{
    // All IExcelRange methods
    // Element-wise iteration yields ExcelScalar elements
    // .Rows iterates row-by-row, .Cols iterates column-by-column
    // For 1×N or N×1 arrays, all three iteration modes produce equivalent results

    // Enables: pConts.Any(c => c == r.Continent)
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

### Value/Cell Access Model

Access to cell formatting (color, bold, etc.) requires COM interop, which is slower and requires `IsMacroType = true`. The wrapper types support **lazy cell escalation**:

- **Default path (fast):** Operations on values. `r.Population > 1000` works via `ColumnValue` operators, no COM needed.
- **Cell path (when needed):** Access `.Cell` on any `ColumnValue` or use `.Cells` on a range. `r.Population.Cell.Color == 6` escalates to COM.
- **Static detection:** Before compilation, the transpiler scans for `.Cell` / `.Cells` usage to determine whether `IsMacroType = true` is needed. Same approach as today's `DetectObjectModelUsage()`.

Shorthand accessors on `IExcelRange`:
- `.Rows` — iterate as `Row` objects (value access + `.Cell` available)
- `.Cols` — iterate as column arrays
- `.Cells` — iterate as `Cell` objects directly (always COM, forces `IsMacroType`)
- Implicit (no accessor) — element-wise on values

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

1. **Parse the expression** — identify the lambda signature and body
2. **Detect inputs:**
   - Single-input sugar: first identifier before `.` is the input
   - Explicit lambda: parameter list defines inputs
   - LET context: additional free variables in the body become inputs
3. **Detect cell usage** — scan for `.Cell` / `.Cells` to set `IsMacroType`
4. **Emit C# code:**
   ```csharp
   public static object UdfName(object input0, object input1, ..., string[] headers0, ...)
   {
       var tblCountries = ExcelValue.Wrap(input0, headers0);  // → ExcelTable
       var maxPop = ExcelValue.Wrap(input1);                  // → ExcelScalar

       // User's code, mostly verbatim
       return tblCountries.Rows.Where(r => r.Population < maxPop).ToResult();
   }
   ```
5. **Result conversion** — `.ToResult()` normalizes output to `object[,]` for Excel

The existing parser (lexer, AST) may need moderate changes to support the explicit lambda syntax. The transpiler's AST-walking code generation is largely replaced by wrapper type method calls.

### Intellisense: Roslyn Completion Service

Replace the custom `CompletionProvider` with Roslyn's `CompletionService`:

1. **Build a synthetic C# document** from the user's code:
   ```csharp
   using System;
   using System.Linq;
   using System.Text.RegularExpressions;
   using FormulaBoss.Runtime;  // pre-compiled wrapper types

   class _Synthetic_ {
       void _M_(ExcelTable tblCountries, ExcelArray pConts, ExcelScalar maxPop) {
           /* user's code here, caret position mapped */
       }
   }
   ```
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

**Performance:** Roslyn completion is designed for interactive use in Visual Studio — fast enough for keystroke-by-keystroke completion. The synthetic document is small (just the user's expression with type declarations), so compilation overhead is minimal.

**Variable type inference for the synthetic document:**
- Table-bound LET variables → `ExcelTable` (with column metadata if known)
- Column-reference LET variables → still passed as header strings (same as today)
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
- [ ] Single-input sugar syntax works: `tblCountries.Rows.Where(r => r.Population > 1000)`
- [ ] Multi-input explicit lambda works: `(tbl, maxPop) => tbl.Rows.Where(r => r.Population < maxPop)`
- [ ] LET variable capture works automatically for scalars, arrays, and tables
- [ ] Nested lambdas resolve correctly: `pConts.Any(c => c == r.Continent)` where `pConts` is an `ExcelArray`
- [ ] Cell escalation works: `r.Population.Cell.Color` triggers `IsMacroType` and returns formatting data
- [ ] Method naming is C# convention everywhere (`.Any()`, `.Where()`, `.All()`, etc.)
- [ ] Roslyn-powered intellisense provides completions for wrapper types AND standard C# methods
- [ ] The target formula from #59 compiles and returns correct results
- [ ] Existing features (LET column bindings, source preservation, error reporting) continue to work

## Out of Scope

- VBA transpiler updates (Phase 12, #21) — will need separate work to understand wrapper types
- Built-in algorithm library (#22) — future work, but wrapper types provide the foundation
- Case-tolerant method names — rely on intellisense to prevent casing errors
- GroupBy detailed design — noted as TBD in the type system, will be designed during implementation

## Design Decisions (Resolved)

1. **ExcelValue.Wrap() factory** — **Runtime type detection.** `Wrap()` inspects the runtime value and returns the appropriate type: `ExcelTable` for ListObject references, `ExcelArray` for `object[,]`, `ExcelScalar` for scalars. This is the safest approach as it handles Excel's inconsistencies (e.g. a formula sometimes returning a scalar, sometimes an array). `ExcelScalar` implements `IExcelRange` with single-element semantics so that calling `.Any()` on a scalar works naturally rather than throwing.

2. **Row dynamic member access** — **Generated TypedRow class** (as today). Concrete typed properties enable Roslyn intellisense for column names at edit time, which is a significant UX win over `DynamicObject` (which would require runtime resolution and provide no completions).

3. **Roslyn workspace management** — **Persistent `AdhocWorkspace`** created at add-in startup, reused across all completion requests. The workspace holds one project with wrapper type references. Each completion request updates only the synthetic document. Memory footprint is small (~few MB). Disposed on shutdown.

4. **ExcelArray shape** — **Always 2D.** `ExcelArray` wraps `object[,]` and preserves its shape. All arrays are conceptually 2D. For a 1×N or N×1 array, element-wise iteration (`.Where()` directly), `.Rows` iteration, and `.Cols` iteration produce equivalent results. `.Select()` flattens to 1D (mapping elements to values). `.Map()` preserves 2D shape.
