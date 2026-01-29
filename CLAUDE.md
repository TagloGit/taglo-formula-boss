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

## UDF Registration Context

UDF registration via `ExcelIntegration.RegisterDelegates` must happen in a valid Excel macro context. When intercepting `SheetChange` events, use `ExcelAsyncUtil.QueueAsMacro()` to defer processing:

```csharp
ExcelAsyncUtil.QueueAsMacro(() =>
{
    // Compile and register UDF here
});
```

## LINQ on object Collections

When working with Excel range values (which are `object[,]`), LINQ aggregations like `.Sum()` don't work directly. Cast values explicitly:

```csharp
// BAD - won't compile
values.Cast<object>().Sum()

// GOOD - cast to numeric type
values.Cast<object>().Select(x => Convert.ToDouble(x)).Sum()
```
