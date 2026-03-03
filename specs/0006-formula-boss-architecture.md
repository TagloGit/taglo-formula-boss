# 0006 — Formula Boss Architecture

How the user-facing behaviour defined in [0005](0005-formula-boss-user-spec.md) is delivered. This document covers the compilation pipeline, type system internals, assembly loading, and runtime infrastructure.

---

## Tech Stack

| Component | Technology | Why |
|---|---|---|
| Excel integration | ExcelDNA 1.9 | Native C API for high-performance UDF registration |
| Runtime compilation | Roslyn (Microsoft.CodeAnalysis.CSharp) | Parse, analyse, and compile C# at runtime |
| Cell formatting | Microsoft.Office.Interop.Excel | COM access for color, bold, font properties |
| Floating editor | AvalonEdit (WPF) | Syntax highlighting, bracket matching, hosted on background STA thread |
| Target framework | .NET 6 / C# 10 | ExcelDNA 1.9 compatibility |

---

## Pipeline Overview

User expression → **Intercept** → **Detect** → **Emit** → **Compile** → **Register** → Excel calls UDF

```
FormulaInterceptor
  → FormulaPipeline.Process()
    → InputDetector.Detect()        // free variable analysis via Roslyn parse
    → CodeEmitter.Emit()            // generate C# UDF source
    → DynamicCompiler.CompileAndRegister()  // Roslyn compile + ALC load + ExcelDNA register
```

### Stage 1: Interception (`FormulaInterceptor`)

- Hooks `Application.SheetChange` event
- Detects backtick formulas via `BacktickExtractor.IsBacktickFormula()`
- Defers processing via `ExcelAsyncUtil.QueueAsMacro()` (UDF registration requires macro context)
- Re-entrancy guard (`_isProcessing`) prevents cascade
- LET formulas handled separately via `ProcessLetFormula` / `LetFormulaParser`

### Stage 2: Detection (`InputDetector`)

Parses the expression with Roslyn to extract metadata:

1. **Range ref preprocessing:** `A1:B10` → `__range_A1_B10` (valid C# identifier)
2. **Roslyn parse:** Wraps expression in `class __Wrapper { object __M() => <expr>; }`
3. **Object model detection:** Scans for `.Cell` / `.Cells` → sets `RequiresObjectModel = true`
4. **Lambda parameter collection:** Extracts all lambda param names to exclude from free vars
5. **Free variable detection:** All identifiers not in (lambda params ∪ C# keywords ∪ type names ∪ local vars ∪ method invocations ∪ `__` prefixes except `__range_`)
6. **Header variable detection:** Traces `r["Col"]` bracket access back to root parameter → marks for header extraction

**Output:** `DetectionResult` — parameters, `RequiresObjectModel`, `HeaderVariables`, normalized expression, range ref map.

### Stage 3: Code Emission (`CodeEmitter`)

Generates a C# class with a static UDF method:

```csharp
using System;
using System.Linq;
using FormulaBoss.Runtime;

public static class __udf_hash_Class
{
    public static object __udf_hash(object param1__raw, object param2__raw, ...)
    {
        var param1 = ExcelValue.Wrap(param1__raw, headers, origin);
        var param2 = ExcelValue.Wrap(param2__raw);

        var __result = /* user expression */;
        return ResultConverter.Convert(__result);
    }
}
```

**UDF naming:** SHA256 hash of expression → first 8 hex chars → `__udf_abcd1234`. Or sanitised preferred name if provided.

**Uniform parameter treatment:** Every parameter gets the same wrapping preamble. The only per-variable distinction is whether to extract headers (based on whether that variable is used with `r["Col"]` bracket access).

### Stage 4: Compilation and Registration (`DynamicCompiler`)

1. **Roslyn compile:** Parse source → create `CSharpCompilation` → emit to `MemoryStream`
2. **ALC loading:** Load into ExcelDNA's `AssemblyLoadContext` (not `Default`) so generated code can resolve `FormulaBoss.Runtime` types
3. **Reflection:** Find public static methods on exported types
4. **Register:** `ExcelIntegration.RegisterDelegates()` with `AllowReference = true` on all params, `IsMacroType` if needed

### Caching

`FormulaPipeline` caches compiled UDFs by expression text. Duplicate expressions reuse existing registrations.

---

## Assembly Loading (ALC)

**The constraint:** Assemblies loaded via `AssemblyLoadContext.Default.LoadFromStream()` cannot resolve types from host-loaded assemblies. The JIT fails silently → `#VALUE!`.

**The solution:** Load generated assemblies into ExcelDNA's named ALC:

```csharp
var alc = AssemblyLoadContext.GetLoadContext(typeof(ExcelFunctionAttribute).Assembly);
var assembly = alc.LoadFromStream(ms);
```

This allows generated code to `using FormulaBoss.Runtime;` and reference all Runtime types directly.

**Requirement:** `FormulaBoss.Runtime.dll` must have `Pack="false"` in the `.dna` file so it's disk-backed (not packed into the XLL).

---

## Delegate Bridge Pattern

For operations that require direct access to ExcelDNA or COM APIs (which cannot be referenced from generated code's context in edge cases), delegate bridges are used:

### RuntimeHelpers (host assembly)

| Delegate | Signature | Purpose |
|---|---|---|
| `ResolveRangeDelegate` | `Func<object, object>` | `XlCall.Excel(xlfReftext, ...)` → COM Range |
| `GetHeadersDelegate` | `Func<object[,], string[]?>` | Extract first row as column headers |
| `GetOriginDelegate` | `Func<object, object?>` | Extract `RangeOrigin` from `ExcelReference` |
| `ToResultDelegate` | `Func<object, object>` | Convert result to `object[,]` for Excel |
| `GetValuesFromReference` | `Func<object, object>` | Extract values from `ExcelReference` |

### RuntimeBridge (Runtime assembly)

| Delegate | Signature | Purpose |
|---|---|---|
| `GetCell` | `Func<string, int, int, Cell>` | Resolve cell at (sheet, row, col) with formatting |

All delegates are initialized in `AddIn.AutoOpen` with lambdas that have direct access to COM/ExcelDNA APIs. Lambdas JIT-compile in the host context where all types resolve.

**Rule:** Bridge classes must never have `using ExcelDna.Integration` or reference any host-loaded type — even in private method bodies.

---

## Type System Internals

### Wrapper Hierarchy

```
ExcelValue (abstract) : IExcelRange, IComparable<ExcelValue>
  ├─ ExcelScalar      — single value, single-element collection semantics
  ├─ ExcelArray       — object[,], always 2D, element-wise iteration
  └─ ExcelTable       — extends ExcelArray, adds string[] Headers and column map
```

### ExcelValue.Wrap() Factory

```csharp
static ExcelValue Wrap(object? value, string[]? headers = null, RangeOrigin? origin = null)
```

- If `headers` provided and `value` is `object[,]` → `ExcelTable`
- If `value` is `object[,]` → `ExcelArray`
- Otherwise → `ExcelScalar`

### Row and Dynamic Dispatch

`Row` extends `DynamicObject`. This enables `r.ColumnName` via `TryGetMember` at runtime. However, CS1977 prevents passing lambdas with `dynamic` params to extension methods. Solution: `RowCollection` provides **instance methods** (not extension methods) accepting `Func<dynamic, ...>`.

### ColumnValue

Returns from `Row[string]` and `Row[int]`. Supports:
- Operator overloading (comparison, arithmetic) with `ColumnValue`, `ExcelValue`, `double`, `int`
- Implicit conversion to `double`, `string`, `bool`
- `.Cell` property for lazy escalation to formatting access (requires `CellAccessor` set by the row's origin context)

### ResultConverter

`ResultConverter.Convert(object? result)` dispatches:
- `ExcelValue` → `.ToResult()` (bare scalar or `object[,]`)
- `IExcelRange` → materialise rows → `object[,]`
- `IEnumerable<Row>` → `object[,]`
- `IEnumerable<ColumnValue>` → `object[,]`
- Primitives → returned as-is

---

## Intellisense Architecture

### Roslyn Workspace (`RoslynWorkspaceManager`)

Persistent `AdhocWorkspace` created at startup, reused across completions. Holds one project with metadata references to `FormulaBoss.Runtime` and standard BCL assemblies.

### Synthetic Document (`SyntheticDocumentBuilder`)

Wraps the user's expression in a valid C# document:

```csharp
using System; using System.Linq; using FormulaBoss.Runtime;
class _Synthetic_ {
    // Synthetic typed Row classes included when table metadata is known
    void _M_(ExcelValue param1, ExcelValue param2, ...) {
        /* user code with caret position mapped */
    }
}
```

When table metadata is known, includes synthetic typed Row classes with real properties for dot-notation completion.

### Completion Flow (`RoslynCompletionProvider`)

1. Build synthetic document from expression + parameters
2. Map editor caret to synthetic document position
3. Call `CompletionService.GetCompletionsAsync()` at mapped position
4. Map Roslyn completions back to `CompletionData` for AvalonEdit
5. Falls back to legacy completions for non-expression contexts (table names, named ranges)

---

## Interception Details

### Standard Backtick Formula

`BacktickExtractor.Extract()` finds all `` ` `` pairs. Each backtick expression is processed independently through the pipeline. The cell formula is rewritten with generated UDF names.

### LET Formula

`LetFormulaParser` parses the LET structure. Each LET binding containing backticks is processed. `LetFormulaRewriter` inserts `_src_varName` documentation bindings to preserve the original source for reconstruction. The final formula is rewritten with `Formula2` for dynamic array support.

---

## IsMacroType Detection

The `InputDetector` scans the Roslyn syntax tree for member access expressions matching `.Cell` or `.Cells`. If found, `RequiresObjectModel = true` flows through the pipeline to `DynamicCompiler`, which sets `IsMacroType = true` on the `ExcelFunctionAttribute` during registration. This allows the UDF to call `xlfReftext` for sheet-qualified addresses needed by COM cell access.

---

## Testing Architecture

Four-tier testing strategy with progressive Excel dependency:

```
Unit Tests → Runtime Tests → Integration Tests → Add-in Tests
(no Excel)     (no Excel)     (Excel COM, no XLL)  (full XLL)
```

### formula-boss.Tests (Unit)

xUnit tests for transpilation components: parser, code emitter, input detector, intellisense, LET formula handling. No Excel dependency.

### formula-boss.Runtime.Tests (Runtime)

xUnit tests for wrapper types (`ExcelValue`, `ExcelArray`, `ExcelScalar`, `ExcelTable`, `Row`, `ColumnValue`, `Cell`) and `ResultConverter`. Pure .NET — no Excel dependency.

### formula-boss.IntegrationTests (Pipeline)

End-to-end pipeline tests: expression → `InputDetector` → `CodeEmitter` → Roslyn compile → execute via reflection. Uses a hidden Excel COM instance for range/cell setup but does NOT load the XLL. Delegate bridges are manually initialized for runtime support.

Key helpers:
- `ExcelTestFixture` — manages hidden Excel instance lifecycle
- `NewPipelineTestHelpers` — compilation and execution infrastructure

### formula-boss.AddinTests (End-to-End)

**The ultimate proof of correctness.** Loads the XLL into a live Excel instance and validates the full pipeline: formula interception → compilation → UDF registration → result display. Tests are serialized via xUnit `[Collection("Excel Addin")]` to prevent Excel contention.

Key helpers:
- `ExcelAddinFixture` (`ICollectionFixture`) — launches hidden Excel, registers XLL, shared across all tests
- `TestUtilities` — COM cell access, formula entry, value polling

**Requirement:** Excel must be installed. Tests assume `formula-boss64.xll` is available in build output.
