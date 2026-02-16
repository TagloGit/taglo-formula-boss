# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Formula Boss is an Excel add-in that allows power users to write inline expressions using a concise DSL that transpiles to C# UDFs at runtime via ExcelDNA and Roslyn. The primary use case is competitive Excel environments.

## Build Commands

```bash
# Build the solution (requires .NET 6 SDK)
dotnet build formula-boss/formula-boss.slnx

# Run tests
dotnet test formula-boss/formula-boss.slnx
```

## Architecture

**Tech Stack:**
- .NET 6 / C# 10 (ExcelDNA 1.9 compatibility)
- ExcelDNA - Native C API integration for high-performance UDFs
- Roslyn (Microsoft.CodeAnalysis.CSharp) - Runtime C# compilation
- Microsoft.Office.Interop.Excel - Excel object model access for cell properties (color, formatting)
- AvalonEdit - Floating editor UI (post-MVP)

## Code Style

Enforced via `.editorconfig` and `Directory.Build.props`:
- File-scoped namespaces
- Private fields: `_camelCase`
- Nullable reference types enabled
- Modern C# patterns preferred (pattern matching, null coalescing, target-typed new)
- Root namespace: `FormulaBoss`

## Warnings

The user monitors warnings using ReSharper in Visual Studio. After significant code refactors, ask the user to review ReSharper warnings before committing.

## Specifications

Detailed specifications and implementation plan are in `docs/specs and plan/`:
- `excel-udf-addin-spec.md` - Full DSL specification, user journeys, error handling
- `excel-udf-addin-implementation-plan.md` - Technical stack rationale, 14 implementation phases, risk register

## ExcelDNA Assembly Identity Issues

**Critical lesson learned:** When dynamically compiling code at runtime with Roslyn that references ExcelDNA types, you WILL encounter assembly identity mismatches.

**The problem:**
- ExcelDNA packs assemblies into the .xll file at build time
- Roslyn compilation references ExcelDNA from the NuGet cache (`~/.nuget/packages/exceldna.integration/`)
- At runtime, objects like `ExcelReference` come from the packed assembly
- Even though type names match (`ExcelDna.Integration.ExcelReference`), .NET treats them as different types
- Pattern matching like `rangeRef is ExcelReference` fails with `#VALUE!` error

**The solution:**
1. **Never reference ExcelDNA types directly in generated code** - use `object` parameter types instead
2. **Use string-based type checking:** `rangeRef?.GetType()?.Name == "ExcelReference"`
3. **Use reflection for method calls:** `rangeRef.GetType().GetMethod("GetValue")?.Invoke(rangeRef, null)`
4. **Use reflection for static API access:** Get `ExcelDnaUtil` via `rangeRef.GetType().Assembly.GetType("ExcelDna.Integration.ExcelDnaUtil")`
5. **Don't generate `[ExcelFunction]` attributes** - register UDFs manually via `ExcelIntegration.RegisterDelegates`
6. **Keep ExcelDna.Integration.dll unpacked** - add `<Reference Path="ExcelDna.Integration.dll" Pack="false" />` in .dna file

**The identity mismatch is transitive.** Generated code also cannot load host assembly types that have ExcelDNA types in their fields, method signatures, or method bodies. Calling `hostAsm.GetType("FormulaBoss.SomeClass")` will throw `TypeLoadException` if `SomeClass` references any ExcelDNA type, because .NET tries to resolve those dependencies in the caller's assembly context.

**Reflection-based XlCall.Excel does not work for C API functions.** Calling `XlCall.Excel(xlfReftext, ...)` via `MethodInfo.Invoke` returns `Object[,]` (range values) instead of the expected address string. Only direct (statically compiled) calls to `XlCall.Excel` work correctly for C API functions like xlfReftext.

**Code patterns to avoid in generated code:**
```csharp
// BAD - will cause assembly identity mismatch
if (rangeRef is ExcelReference excelRef)
using ExcelDna.Integration;
[ExcelFunction(...)]

// GOOD - use reflection
if (rangeRef?.GetType()?.Name == "ExcelReference")
var getValueMethod = rangeRef.GetType().GetMethod("GetValue");
```

## Delegate Bridge Pattern

When generated code needs to call ExcelDNA C API functions (like `xlfReftext`), use a delegate bridge:

1. **Define a `Func<>` field** on a host class with NO ExcelDNA dependencies (e.g. `RuntimeHelpers.ResolveRangeDelegate`)
2. **Initialize it from `AddIn.AutoOpen`** with a lambda that calls `XlCall.Excel(...)` directly
3. **Generated code invokes the delegate** by finding the field via reflection on the host assembly

The lambda is JIT-compiled in the host context (where ExcelDNA resolves), but the field type (`Func<object, object>`) has no ExcelDNA dependency, so the host class can be loaded from generated code. See `RuntimeHelpers.ResolveRangeDelegate` and `AddIn.AutoOpen` for the implementation.

**Critical rule:** `RuntimeHelpers` (and any class loaded from generated code) must NEVER have `using ExcelDna.Integration` or reference any ExcelDNA type — not even in private method bodies. Any such reference causes `TypeLoadException` when the type is loaded from the Roslyn-compiled assembly's context.

## Object Model UDFs and IsMacroType

UDFs that access cell formatting (color, bold, etc.) via the `.cells` accessor require `IsMacroType = true` in their `ExcelFunctionAttribute`. This is because `xlfReftext` (needed for sheet-qualified range addresses) is a C API function that only works from macro-type UDFs. The `RequiresObjectModel` flag on `TranspileResult` controls this — it flows through `FormulaPipeline` into `DynamicCompiler.CompileAndRegister`.

## UDF Registration Context

UDF registration via `ExcelIntegration.RegisterDelegates` must happen in a valid Excel macro context. When intercepting `SheetChange` events, use `ExcelAsyncUtil.QueueAsMacro()` to defer processing:

```csharp
ExcelAsyncUtil.QueueAsMacro(() =>
{
    // Compile and register UDF here
});
```

## Formula2 for Dynamic Array Spilling

When programmatically setting cell formulas that return arrays, use `Formula2` instead of `Formula`:

```csharp
// BAD - adds implicit intersection @ operator, prevents spilling
cell.Formula = "=MyUDF(A1:A5)";

// GOOD - enables dynamic array spilling
cell.Formula2 = "=MyUDF(A1:A5)";
```

The `Formula` property is legacy and Excel will add an `@` prefix to prevent spilling. `Formula2` is the dynamic array-aware property introduced in Excel 365.

## LINQ on object Collections

When working with Excel range values (which are `object[,]`), LINQ aggregations like `.Sum()` don't work directly. Cast values explicitly:

```csharp
// BAD - won't compile
values.Cast<object>().Sum()

// GOOD - cast to numeric type
values.Cast<object>().Select(x => Convert.ToDouble(x)).Sum()
```
