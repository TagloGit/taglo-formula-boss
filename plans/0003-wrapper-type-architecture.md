# 0003 — Wrapper-Type Architecture — Implementation Plan

## Overview

Replace the current DSL syntax transformation with a pre-compiled type-safe wrapper library. Formula Boss expressions become standard C# lambdas operating on typed facades (`ExcelTable`, `ExcelArray`, `ExcelScalar`). The custom lexer/parser is deleted; Roslyn handles all C# parsing, input detection, compilation, and intellisense.

**Spec:** `specs/0003-wrapper-type-architecture.md`
**Epic:** #59

## Key Decisions

- **Incremental PRs** — each phase is a separate mergeable PR
- **Separate project** — `FormulaBoss.Runtime` (.csproj) for wrapper types, clean assembly boundary
- **Roslyn for input detection** — delete custom lexer/parser, use `CSharpSyntaxTree` for everything
- **COM access strategy** — spike early to determine if wrapper assembly can reference Office interop directly or needs delegate bridges

## Phase 1: Runtime Assembly Spike + Core Value Types

**Goal:** Create `FormulaBoss.Runtime` project, spike the assembly identity question, implement core value types without COM access.

**Files to create:**
- `formula-boss.Runtime/FormulaBoss.Runtime.csproj` — net6.0-windows, no ExcelDNA reference
- `formula-boss.Runtime/ExcelValue.cs` — base class + `Wrap()` factory
- `formula-boss.Runtime/ExcelScalar.cs` — single value wrapper with operators and `IExcelRange` single-element semantics
- `formula-boss.Runtime/ExcelArray.cs` — `object[,]` wrapper with `IExcelRange` implementation
- `formula-boss.Runtime/ExcelTable.cs` — table wrapper (ListObject-backed), `IExcelRange`
- `formula-boss.Runtime/IExcelRange.cs` — shared interface
- `formula-boss.Runtime/Row.cs` — row access with bracket/index notation
- `formula-boss.Runtime/ColumnValue.cs` — column value with implicit conversions and operators
- `formula-boss.Runtime/ResultConverter.cs` — `.ToResult()` extension, normalises output to `object[,]`

**Files to create (tests):**
- `formula-boss.Runtime.Tests/FormulaBoss.Runtime.Tests.csproj`
- `formula-boss.Runtime.Tests/ExcelScalarTests.cs`
- `formula-boss.Runtime.Tests/ExcelArrayTests.cs`
- `formula-boss.Runtime.Tests/ExcelTableTests.cs`
- `formula-boss.Runtime.Tests/RowTests.cs`
- `formula-boss.Runtime.Tests/ColumnValueTests.cs`
- `formula-boss.Runtime.Tests/ResultConverterTests.cs`

**Spike:** Add `Microsoft.Office.Interop.Excel` reference to Runtime project. Write a test that loads the Runtime assembly from a simulated "generated code" context (separate `AssemblyLoadContext`) and checks whether Office interop types resolve. This determines whether we need delegate bridges for COM access or can reference interop directly.

**What this PR delivers:**
- All value-path operations working: `Where`, `Select`, `Any`, `All`, `First`, `FirstOrDefault`, `OrderBy`, `OrderByDescending`, `Take`, `Skip`, `Distinct`, `Count`, `Sum`, `Min`, `Max`, `Average`, `Aggregate`, `Scan`, `Map`, `SelectMany`
- `ExcelValue.Wrap()` factory with runtime type detection
- `ColumnValue` with implicit conversions (`double`, `string`, `bool`) and comparison/arithmetic operators
- `Row` with bracket access (`r[0]`, `r["ColName"]`) and dynamic member access for dot notation
- `ExcelScalar` implementing `IExcelRange` with single-element semantics
- Result conversion to `object[,]`
- Assembly identity spike result documented

**Not in this PR:** Cell/COM access, transpiler changes, intellisense.

## Phase 2: Cell Escalation and COM Access

**Goal:** Add `Cell`, `Interior`, `Font` types and the `.Cell` property on `ColumnValue` / `.Cells` accessor on `IExcelRange`.

**Depends on Phase 1 spike result:**
- **If interop loads cleanly:** Wrapper types reference `Microsoft.Office.Interop.Excel` directly. `Cell` wraps a COM `Range` object.
- **If identity mismatch:** Define `Func<>` delegate fields on a static class in Runtime (e.g. `RuntimeBridge`). Host initialises them at startup with lambdas that do the actual COM calls. `Cell` invokes delegates.

**Files to create/modify:**
- `formula-boss.Runtime/Cell.cs` — cell wrapper with properties (`Color`, `Rgb`, `Bold`, `Italic`, `FontSize`, `Format`, `Formula`, `Row`, `Col`, `Address`, `Value`, `Interior`, `Font`)
- `formula-boss.Runtime/Interior.cs` — Interior sub-object wrapper
- `formula-boss.Runtime/Font.cs` — Font sub-object wrapper
- `formula-boss.Runtime/ColumnValue.cs` — add `.Cell` property
- `formula-boss.Runtime/IExcelRange.cs` — add `.Cells` accessor
- `formula-boss.Runtime/ExcelArray.cs` — implement `.Cells` iteration
- `formula-boss.Runtime/ExcelTable.cs` — implement `.Cells` iteration

**Files to create (tests):**
- `formula-boss.Runtime.Tests/CellTests.cs`
- `formula-boss.Runtime.Tests/CellEscalationTests.cs`

**If delegate bridge needed:**
- `formula-boss.Runtime/RuntimeBridge.cs` — static delegate fields for COM operations
- `formula-boss/AddIn.cs` — initialise bridge delegates in `AutoOpen()`

**What this PR delivers:**
- `r.Price.Cell.Color` works (cell escalation from ColumnValue)
- `.Cells` accessor iterates as `Cell` objects
- All `Cell` properties from the spec
- `Interior` and `Font` sub-objects

## Phase 3: Transpiler Rewrite

**Goal:** Replace the DSL transpiler with a Roslyn-based input detector and thin code emitter. Delete the custom lexer, parser, and old transpiler.

> **Lesson from first attempt (Feb 2026):** Generated code cannot reference `FormulaBoss.Runtime` types directly due to the assembly identity constraint (see `docs/assembly-identity-investigation.md`). The CodeEmitter must use delegate bridges for all Runtime type interactions. Additionally: migrate relevant old test cases before deleting old code, and include integration tests that compile and execute generated code.

### Step 1: Migrate test cases (before deleting anything)

Review the existing test files and extract test cases that cover edge cases the new implementation must handle:

- `formula-boss.Tests/TranspilerTests.cs` — extract cases for: range ref parsing (`A1:A10`, `$A$1:$B$10`), UDF naming (hash-based `__udf_` prefix, collision avoidance, reserved Excel names), `IsMacroType` detection, LET variable capture, multi-input handling
- `formula-boss.Tests/ParserTests.cs` — extract cases for: operator precedence, member access chains, lambda parameter handling, bracket notation (`r[0]`, `r["Col Name"]`)

Create `formula-boss.Tests/TranspilerMigrationCases.md` documenting each extracted case and what the new test should verify. This is a checklist, not test code.

### Step 2: Build InputDetector and CodeEmitter

**Files to create:**
- `formula-boss/Transpilation/InputDetector.cs` — uses Roslyn `CSharpSyntaxTree.ParseText()` to:
  - Detect single-input sugar vs explicit lambda
  - Extract input identifier names
  - Detect `.Cell`/`.Cells` usage for `IsMacroType`
  - Detect free variables (for LET capture)
  - Handle range references (`A1:A10`) — these are not valid C# identifiers, so the detector must handle them as special tokens or pre-process them
- `formula-boss/Transpilation/CodeEmitter.cs` — generates the UDF method:
  - **No direct references to `FormulaBoss.Runtime` types** in generated code
  - Uses delegate bridges on `RuntimeHelpers` for: wrapping inputs (`WrapDelegate`), converting results (`ToResultDelegate`), getting headers/origin
  - All UDF parameters are `object` type
  - Return type is `object` (not `object?[,]`)
  - No namespace wrapper in generated code
  - UDF names use `__udf_` + hash prefix (never derived from input parameter names)
  - User's expression/lambda body passed through verbatim

**Delegate bridges to add on `RuntimeHelpers`:**
- `WrapDelegate: Func<object, string[], object, object>` — calls `ExcelValue.Wrap(value, headers, origin)`, returns `object` (the wrapped ExcelValue)
- `ToResultDelegate: Func<object, object>` — calls `ResultConverter.ToResult(result)`, returns `object[,]`
- `GetHeadersDelegate: Func<object, string[]>` — calls through to header extraction
- `GetOriginDelegate: Func<object, object>` — calls through to origin extraction
- `GetValuesDelegate: Func<object, object>` — calls through to value extraction from ExcelReference

These are initialized in `AddIn.AutoOpen` with lambdas that reference Runtime types directly (host context).

**Generated code pattern:**
```csharp
// NO namespace, NO Runtime using statements
public static class __udf_A1B2C3D4_Class
{
    public static object __udf_A1B2C3D4(object tbl__raw)
    {
        // Resolve ExcelReference to values
        var tbl__isRef = tbl__raw?.GetType()?.Name == "ExcelReference";
        var tbl__values = tbl__isRef == true
            ? FormulaBoss.RuntimeHelpers.GetValuesDelegate(tbl__raw)
            : tbl__raw;
        var tbl__headers = tbl__isRef == true
            ? FormulaBoss.RuntimeHelpers.GetHeadersDelegate(tbl__raw)
            : null;
        var tbl__origin = tbl__isRef == true
            ? FormulaBoss.RuntimeHelpers.GetOriginDelegate(tbl__raw)
            : null;

        // Wrap via delegate bridge (returns object, actually ExcelValue at runtime)
        dynamic tbl = FormulaBoss.RuntimeHelpers.WrapDelegate(tbl__values, tbl__headers, tbl__origin);

        // User's code — operates on dynamic, Runtime types resolve at runtime
        object __result = tbl.Rows.Where(r => r["Unit Price"] > 5);

        // Convert result via delegate bridge
        return FormulaBoss.RuntimeHelpers.ToResultDelegate(__result);
    }
}
```

**Key difference from first attempt:** Generated code uses `dynamic` for wrapped values and delegate bridges for all Runtime interactions. The `dynamic` keyword means method calls like `.Rows`, `.Where()` resolve at runtime via the DLR, where the actual `ExcelValue` types are available. No compile-time type references needed.

### Step 3: Wire pipeline and update compiler

**Files to modify:**
- `formula-boss/Interception/FormulaPipeline.cs` — replace lexer→parser→transpiler with `InputDetector` → `CodeEmitter` → `DynamicCompiler`
- `formula-boss/Compilation/DynamicCompiler.cs` — Runtime `MetadataReference` is no longer needed for generated code (generated code doesn't reference Runtime types). Keep it only if other compilation needs require it.
- `formula-boss/RuntimeHelpers.cs` — add `WrapDelegate`, `ToResultDelegate`, `GetHeadersDelegate`, `GetOriginDelegate`, `GetValuesDelegate` fields
- `formula-boss/AddIn.cs` — initialize new delegate bridges in `AutoOpen`
- `formula-boss/Transpilation/TranspileResult.cs` — simplify if needed

### Step 4: Delete old code

Only after all new tests pass:

- `formula-boss/Parsing/Lexer.cs`
- `formula-boss/Parsing/Parser.cs`
- `formula-boss/Parsing/Ast.cs`
- `formula-boss/Transpilation/CSharpTranspiler.cs`
- `formula-boss/Transpilation/ExcelTypeSystem.cs`
- `formula-boss.Tests/TranspilerTests.cs`
- `formula-boss.Tests/ParserTests.cs`

Note: `Lexer.cs`, `Parser.cs`, `Ast.cs`, `ExcelTypeSystem.cs` are still used by `CompletionProvider` and `ErrorHighlighter`. Keep them if Phase 5 hasn't landed yet; delete them in Phase 5 or 6 when those UI components are rewritten.

### Tests required in this PR

**Unit tests:**
- `formula-boss.Tests/InputDetectorTests.cs`:
  - Sugar syntax: `tbl.Rows.Where(...)` → primary input `tbl`
  - Explicit lambda: `(a, b) => a.Rows.Count()` → inputs `[a, b]`
  - Statement block: `(tbl) => { return tbl.Rows.Count(); }` → input `tbl`, `IsStatementBody = true`
  - Range refs: `A1:A10.Rows.Count()` → must not truncate to `A1`
  - Object model detection: `.Cell`, `.Cells` → `RequiresObjectModel = true`
  - Free variable detection: `(tbl) => tbl.Rows.Where(r => r.X < maxVal)` → free var `maxVal`
  - String bracket detection: `r["Col Name"]` → detected as requiring headers
- `formula-boss.Tests/CodeEmitterTests.cs`:
  - Generated code contains delegate bridge calls, not direct Runtime references
  - Generated code uses `dynamic` for wrapped values
  - UDF naming: `__udf_` + hash prefix, never matches input parameter name
  - Return type is `object`
  - No namespace in generated code
  - Free variables become additional `object` parameters with wrapping
  - `RequiresObjectModel` propagated to `TranspileResult`

**Integration tests (compile and execute):**
- `formula-boss.IntegrationTests/WrapperTypePipelineTests.cs`:
  - Expression → InputDetector → CodeEmitter → Roslyn compile → load assembly → invoke method → verify result
  - Tests must exercise the Runtime types at runtime (the compiled code uses `dynamic`, so the Runtime types resolve when the test assembly loads them)
  - Minimum scenarios:
    - Single-input value filtering: `tbl.Where(v => v > 5)` (element-wise)
    - Row filtering: `tbl.Rows.Where(r => r[0] > 10)`
    - Row filtering with string key: `tbl.Rows.Where(r => r["Price"] > 10)` (requires headers)
    - Multi-input: `(tbl, maxVal) => tbl.Rows.Where(r => r[0] < maxVal)`
    - Aggregation: `tbl.Rows.Count()`
    - Statement block: `(tbl) => { var c = tbl.Rows.Count(); return c; }`
    - Range ref in sugar syntax (verify `A1:A10` not truncated)
    - UDF name doesn't collide with parameter name

**What this PR delivers:**
- Single-input sugar: `` `tbl.Rows.Where(r => r.X > 0)` `` works
- Explicit lambda: `` `(tbl, maxPop) => tbl.Rows.Where(r => r.X < maxPop)` `` works
- Statement blocks: `` `(tbl) => { return tbl.Rows.Count(); }` `` works
- LET variable capture: free variables detected and added as UDF parameters
- `.Cell`/`.Cells` detection sets `IsMacroType`
- All Runtime type interaction via delegate bridges — no direct type references in generated code
- Old lexer/parser/transpiler deleted (unless still needed by UI components)
- All method names are C# convention (`.Any()`, `.Where()`, etc.)
- Comprehensive test coverage including integration tests that compile and execute generated code

## Phase 4: TypedRow Generation and Column Intellisense Support

**Goal:** Generate `TypedRow` classes with concrete properties for known columns, enabling Roslyn intellisense for `r.Population` etc.

**Files to create/modify:**
- `formula-boss/Transpilation/TypedRowGenerator.cs` — given column names, generates a class inheriting from `Row` with typed properties returning `ColumnValue`
- `formula-boss/Transpilation/CodeEmitter.cs` — when column metadata is available, emit `TypedRow` class and use it in generated code
- `formula-boss.Runtime/Row.cs` — ensure it's designed for inheritance (virtual/protected as needed)

**Files to create (tests):**
- `formula-boss.Tests/TypedRowGeneratorTests.cs`

**What this PR delivers:**
- `r.Population`, `r.Continent` etc. work with Roslyn providing completions
- Bracket access (`r["Column Name"]`, `r[0]`) continues to work as fallback
- Column metadata flows from LET bindings through to TypedRow generation

## Phase 5: Roslyn Intellisense

**Goal:** Replace the custom `CompletionProvider` with Roslyn's `CompletionService` operating on a synthetic C# document.

**Files to delete:**
- `formula-boss/UI/CompletionProvider.cs` (or heavily rewrite)
- `formula-boss.Tests/CompletionScopingTests.cs` (replace with new tests)

**Files to create:**
- `formula-boss/UI/RoslynCompletionProvider.cs` — builds synthetic document, calls `CompletionService.GetCompletionsAsync()`, maps results back
- `formula-boss/UI/SyntheticDocumentBuilder.cs` — constructs the synthetic C# document with wrapper type variable declarations from context
- `formula-boss/UI/RoslynWorkspaceManager.cs` — persistent `AdhocWorkspace` lifecycle, project with Runtime assembly reference

**Files to modify:**
- `formula-boss/UI/FloatingEditorWindow.xaml.cs` — wire up new completion provider
- `formula-boss/UI/CompletionData.cs` — may need adaptation for Roslyn completion items

**NuGet additions:**
- `Microsoft.CodeAnalysis.Features` (or `Microsoft.CodeAnalysis.CSharp.Features`) — for `CompletionService`
- `Microsoft.CodeAnalysis.Workspaces.Common` — for `AdhocWorkspace`

**Files to create (tests):**
- `formula-boss.Tests/RoslynCompletionTests.cs` — test completions for wrapper types, string methods, LINQ
- `formula-boss.Tests/SyntheticDocumentTests.cs` — test document construction from various contexts

**What this PR delivers:**
- Completions for all wrapper type methods and properties
- Completions for C# string methods, LINQ, regex
- Column name completions on Row types (from TypedRow)
- Context-aware: knows type at caret position
- Persistent workspace, fast per-keystroke performance

## Phase 6: Syntax Highlighting and Error Diagnostics Update

**Goal:** Update the floating editor's syntax highlighting for the new C# method names and integrate Roslyn diagnostics for real-time error squiggles.

**Files to modify:**
- `formula-boss/UI/SyntaxHighlighting.xshd` (or equivalent) — update method/property names to PascalCase C# convention
- `formula-boss/UI/ErrorHighlighter.cs` — integrate Roslyn diagnostics instead of custom parser errors
- `formula-boss/UI/FloatingEditorWindow.xaml.cs` — wire up real-time Roslyn diagnostics (debounced)

**What this PR delivers:**
- Syntax highlighting matches new API (`.Where`, `.Any`, `.Cells`, etc.)
- Real-time Roslyn compile errors shown as squiggles
- Error messages are standard Roslyn/C# errors (more familiar to users)

## Phase 7: Integration Testing and Cleanup

**Goal:** End-to-end validation against the target formula from #59, cleanup dead code, update documentation.

**Files to modify:**
- `formula-boss.IntegrationTests/` — add integration tests for the full pipeline with wrapper types
- `specs/0001-excel-udf-addin.md` — final pass to ensure consistency with implementation
- `CLAUDE.md` — update architecture section to reflect wrapper types

**Files to potentially delete:**
- Any remaining dead code from the old DSL system
- `formula-boss/Transpilation/ExcelTypeSystem.cs` (if not already deleted in Phase 3)

**Acceptance criteria validation:**
- [ ] Target formula from #59 compiles and returns correct results
- [ ] Single-input sugar syntax works
- [ ] Multi-input explicit lambda works
- [ ] LET variable capture works for scalars, arrays, tables
- [ ] Nested lambdas resolve correctly
- [ ] Cell escalation triggers IsMacroType
- [ ] C# convention method naming everywhere
- [ ] Roslyn intellisense for wrapper types AND standard C#
- [ ] Existing features (LET column bindings, source preservation, error reporting) work

## Testing Approach

Each phase includes its own tests:

- **Runtime types (Phases 1–2):** Pure unit tests in `formula-boss.Runtime.Tests`. No Excel dependency. Test operators, conversions, collection operations, cell access.
- **Transpiler (Phases 3–4):** Unit tests for input detection and code emission. Integration tests that compile generated code and verify it runs.
- **Intellisense (Phase 5):** Unit tests that verify completion results for various cursor positions and contexts.
- **End-to-end (Phase 7):** Integration tests with real wrapper types + compiled UDFs. The target formula from #59 is the primary acceptance test.

## Risks

| Risk | Mitigation |
|------|------------|
| Assembly identity with Office interop in Runtime assembly | Phase 1 spike tests this explicitly before building on it |
| Roslyn `CompletionService` NuGet package size/compatibility | Phase 5 checks package availability for net6.0; fallback is completion from pre-built lists (similar to today) |
| Performance of Roslyn parsing for input detection (on every formula) | Roslyn `SyntaxTree.ParseText()` is ~1ms for small snippets; cache if needed |
| Breaking change — all existing formulas stop working | Pre-release product, accepted in spec. No migration path. |
| `DynamicObject` vs TypedRow for column access | Phase 4 generates TypedRow for known columns; `DynamicObject` fallback for unknown columns |

## Open Questions

None — all resolved during planning. The Phase 1 spike will finalise the COM access approach.
