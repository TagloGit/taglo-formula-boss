# Cross-Sheet Cell Property Access Bug

## Problem

Cell property access (`.color`, `.bold`, etc.) via the `.cells` accessor does not work correctly with cross-sheet references.

Example: `someNamedRange.cells.select(c => c.color)` where `someNamedRange` is defined on Sheet1.

## Symptoms

- **Same sheet**: Returns correct values (e.g. ColorIndex=6 for yellow cells)
- **Different sheet**: Returns -4142 (`xlNone`) for all cells
- **Recalculation dependency on active sheet**: When F9 is pressed from Sheet1, both sheets show correct results. When F9 is pressed from Sheet2, BOTH sheets show -4142 — even the formula on Sheet1. This indicates the active sheet at recalculation time affects all UDF calls.

## Root Cause Analysis

### How the generated code resolves ranges

The generated UDF code (object model path in `CSharpTranspiler.cs`) converts an `ExcelReference` to a COM `Range` by:

1. Reading `RowFirst`, `RowLast`, `ColumnFirst`, `ColumnLast` properties via reflection
2. Building an A1-style address string (e.g. `A1:C3`) — **without any sheet qualifier**
3. Calling `app.Range["A1:C3"]` which resolves against the **active sheet**

This means all cell property reads hit whatever sheet is active during recalculation, not the sheet the reference actually points to.

### Why xlfReftext was not used originally

The generated code runs in a Roslyn-compiled assembly. Due to ExcelDNA assembly identity mismatches (see CLAUDE.md), the generated code uses reflection for all ExcelDNA API access. The reflection-based call to `XlCall.Excel(xlfReftext, ref, true)` does not work correctly — it returns an `Object[,]` (the range values) instead of the address string.

### What does work

Diagnostic testing confirmed:
- **Direct `XlCall.Excel(XlCall.xlfReftext, ref, true)` from the host assembly** returns the correct sheet-qualified address (e.g. `[Book1]Sheet1!$A$1:$C$3`) from both same-sheet and cross-sheet contexts
- **`Application.Range` with a sheet-qualified address** resolves to the correct sheet
- **`Interior.ColorIndex` on cross-sheet Range objects** returns correct values (confirmed via diagnostic UDF `FB_DiagRef`)
- **`IsMacroType = true`** on the ExcelFunctionAttribute is required for xlfReftext to work (non-macro-type UDFs cannot call C API functions like xlfReftext)

### Key constraint

The generated code (Roslyn-compiled) cannot directly reference ExcelDNA types due to assembly identity mismatches. But `RuntimeHelpers` in the host assembly CAN use ExcelDNA types directly. The challenge is bridging these two contexts for range resolution.

## Diagnostic UDF

A diagnostic function `FB_DiagRef` was used to test each step of range resolution independently (type check, SheetId, xlfReftext, Range resolution, Worksheet name, ColorIndex). This confirmed the address resolution is the root cause, not the COM property access itself. The diagnostic code is not committed but can be recreated from the description in this document.

## Files Involved

- `formula-boss/Transpilation/CSharpTranspiler.cs` — generates the object model range resolution code (~line 1680+)
- `formula-boss/RuntimeHelpers.cs` — host-assembly helper that CAN use XlCall directly
- `formula-boss/Compilation/DynamicCompiler.cs` — UDF registration (needs `IsMacroType = true` for object model UDFs)

## Requirements for a Fix

1. Object model UDFs must be registered with `IsMacroType = true` so xlfReftext works
2. Range resolution must produce a sheet-qualified address
3. The solution must work from Roslyn-generated code despite ExcelDNA assembly identity constraints
4. COM Range objects from cross-sheet references must correctly expose formatting properties

## Resolution

Fixed via the **delegate bridge pattern**. Three constraints interact:

1. **Generated code can't reference ExcelDNA types** (assembly identity mismatch — documented in CLAUDE.md)
2. **Generated code can't load host types that reference ExcelDNA** (transitive `TypeLoadException` — even `using ExcelDna.Integration` in a method body makes the whole class unloadable)
3. **`XlCall.Excel(xlfReftext)` only works via direct (statically compiled) calls** — reflection-based invocation returns `Object[,]` instead of the address string

The solution uses a `Func<object, object>` delegate stored on `RuntimeHelpers` (which has zero ExcelDNA dependencies). `AddIn.AutoOpen` initializes this delegate with a lambda that calls `XlCall.Excel(xlfReftext, rangeRef, true)` directly. The lambda is JIT-compiled in the host context where ExcelDNA types resolve. Generated code finds `RuntimeHelpers` via `AppDomain.GetAssemblies()` reflection and invokes `GetRangeFromReference`, which calls the delegate.

### Files changed

- `formula-boss/AddIn.cs` — initializes `RuntimeHelpers.ResolveRangeDelegate` with direct `XlCall.Excel(xlfReftext)` call
- `formula-boss/RuntimeHelpers.cs` — `GetRangeFromReference` delegates to the bridge; no ExcelDNA types anywhere
- `formula-boss/Transpilation/CSharpTranspiler.cs` — generated code calls `RuntimeHelpers` via reflection instead of inline address building
- `formula-boss/Compilation/DynamicCompiler.cs` — `IsMacroType = true` for object model UDFs
- `formula-boss/Interception/FormulaPipeline.cs` — passes `RequiresObjectModel` through to compiler
