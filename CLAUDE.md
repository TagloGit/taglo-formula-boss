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

Detailed specifications and implementation plans:
- `specs/0001-excel-udf-addin.md` - Full DSL specification, user journeys, error handling
- `plans/0001-excel-udf-addin.md` - Technical stack rationale, 14 implementation phases, risk register

## Workflow and Skills

This repo follows the [Taglo Claude Code workflow](https://github.com/TagloGit/taglo-pm/blob/main/docs/claude-code-workflow.md).

### Available Skills and Agents
- `/developer <issue>` — Execute implementation work
- `/qa <pr-or-issue>` — Test coverage review
- `/planner <topic>` — Spec and plan creation (global skill)
- `/pm` — Cross-repo status (global skill)
- `/backlog-admin` — Quick issue creation (global skill)
- Reviewer subagent — PR review (`.claude/agents/reviewer.md`)

### Specs and Plans
- Specs: `specs/NNNN-short-title.md`
- Plans: `plans/NNNN-short-title.md`

## Assembly Identity — Generated Code Cannot Reference Host Types

**Critical constraint:** Roslyn-compiled code loaded via `AssemblyLoadContext.Default.LoadFromStream()` cannot directly reference types from ANY assembly loaded by the host process. This is a .NET runtime limitation, not ExcelDNA-specific. It affects `ExcelDna.Integration`, `FormulaBoss.Runtime`, and any other host-loaded assembly.

**Root cause:** `LoadFromStream()` creates an assembly without file-backed identity. The JIT cannot match type references in the stream-loaded assembly to already-loaded assemblies, even when Roslyn compiled against the exact same DLL. The JIT fails silently; ExcelDNA surfaces it as `#VALUE!`.

**What does NOT fix this:** `Pack="false"` in the `.dna` file, pointing Roslyn at the loaded assembly's `Location`, or using `MetadataReference.CreateFromImage()`. See `docs/assembly-identity-investigation.md` for the full investigation and test results.

**The rules for generated code:**
1. **Never reference any host type directly** — no `typeof()`, `using`, or direct method calls to `ExcelDna.*`, `FormulaBoss.Runtime.*`, etc.
2. **Use `object` parameter types** and string-based type checking: `obj?.GetType()?.Name == "ExcelReference"`
3. **Use reflection for one-off calls:** `obj.GetType().GetMethod("Foo")?.Invoke(obj, args)`
4. **Use the delegate bridge pattern** for repeated or complex operations (see below)
5. **Don't generate `[ExcelFunction]` attributes** — register UDFs manually via `ExcelIntegration.RegisterDelegates`

**The identity mismatch is transitive.** Generated code cannot load host types that reference other host types in their fields, method signatures, or method bodies. `hostAsm.GetType("FormulaBoss.SomeClass")` will throw `TypeLoadException` if `SomeClass` references any ExcelDNA or Runtime type.

**Reflection-based XlCall.Excel does not work for C API functions.** `XlCall.Excel(xlfReftext, ...)` via `MethodInfo.Invoke` returns `Object[,]` instead of the expected address string. Only direct (statically compiled) calls work. This is why the delegate bridge pattern exists.

**Code patterns:**
```csharp
// BAD — will cause #VALUE! at runtime (any host assembly, not just ExcelDNA)
typeof(FormulaBoss.Runtime.ExcelValue)
FormulaBoss.Runtime.ExcelValue.Wrap(x)
if (rangeRef is ExcelReference excelRef)
using ExcelDna.Integration;
using FormulaBoss.Runtime;

// GOOD — reflection or delegate bridges
if (rangeRef?.GetType()?.Name == "ExcelReference")
var getValueMethod = rangeRef.GetType().GetMethod("GetValue");
RuntimeHelpers.WrapDelegate?.Invoke(rawValue)  // delegate bridge
```

## Delegate Bridge Pattern

When generated code needs to interact with host-loaded assemblies (ExcelDNA, FormulaBoss.Runtime), use delegate bridges:

1. **Define `Func<>` / `Action<>` fields** on a bridge class whose signatures use ONLY primitive types and `object` — no types from ExcelDNA or FormulaBoss.Runtime
2. **Initialize delegates from `AddIn.AutoOpen`** with lambdas that call the actual typed APIs (lambdas JIT-compile in the host context where all types resolve)
3. **Generated code invokes the delegates** — it can load the bridge class because the class has no problematic type dependencies

**Why this works:** The delegate fields use types like `Func<object, object>` from the base class library — no cross-assembly resolution needed. The lambdas are compiled in the host context where all types are known.

**Current bridge classes:**
- `RuntimeHelpers` — delegates for ExcelDNA operations (`ResolveRangeDelegate`, `GetValuesFromReference`, etc.)
- `RuntimeBridge` (in FormulaBoss.Runtime) — delegates for COM/cell access (`GetCell`, `GetHeaders`, `GetOrigin`)

**Critical rule:** Bridge classes must NEVER have `using ExcelDna.Integration`, `using FormulaBoss.Runtime`, or reference any host-loaded type — not even in private method bodies. Any such reference causes `TypeLoadException` when the class is loaded from generated code's context.

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

## COM Object Lifecycle and Excel Shutdown

**Every COM object obtained from Excel interop must be explicitly released via `Marshal.ReleaseComObject()`.** Unreleased COM references can prevent Excel from shutting down cleanly, leaving zombie processes.

**Rules:**
- Release transient COM objects (`ActiveCell`, `ActiveWindow`, `Range`, `Worksheet`) in `finally` blocks after use
- Release stored COM references before overwriting them with new values
- Capture COM objects on the Excel thread, not inside WPF dispatcher lambdas (avoids cross-apartment proxy issues)
- On shutdown, release all stored COM references BEFORE shutting down background threads

**ExcelDNA shutdown:** `AutoClose` is NOT called when Excel shuts down — only when the add-in is explicitly removed. Use `ExcelComAddIn.OnBeginShutdown` to detect Excel closing (see `AddIn.ShutdownMonitor`). Reference: https://excel-dna.net/docs/guides-advanced/detecting-excel-shutdown-and-autoclose/

**WPF thread cleanup:** `Dispatcher.InvokeShutdown()` is async — always `Thread.Join()` afterward to ensure the thread exits before continuing cleanup.

## LINQ on object Collections

When working with Excel range values (which are `object[,]`), LINQ aggregations like `.Sum()` don't work directly. Cast values explicitly:

```csharp
// BAD - won't compile
values.Cast<object>().Sum()

// GOOD - cast to numeric type
values.Cast<object>().Select(x => Convert.ToDouble(x)).Sum()
```
