# Assembly Identity Mismatch — Investigation and Findings

## Date

February 2026

## Context

During implementation of the wrapper-type architecture (#59, #62), we discovered that Roslyn-compiled UDFs cannot directly reference types from `FormulaBoss.Runtime` — the same constraint previously known for `ExcelDna.Integration`. This investigation determined whether `Pack="false"` in the `.dna` file could solve the problem.

## Background

Formula Boss compiles C# code at runtime using Roslyn and loads it via `AssemblyLoadContext.Default.LoadFromStream()`. The generated assemblies need to interact with types from:

- `ExcelDna.Integration` — for `ExcelReference`, `XlCall`, etc.
- `FormulaBoss.Runtime` — for wrapper types (`ExcelValue`, `ExcelArray`, `Row`, etc.)
- `formula-boss.dll` — for `RuntimeHelpers` (must have no ExcelDNA/Runtime type references)

The existing workaround (delegate bridge pattern) uses `Func<object, object>` fields on classes with no problematic type dependencies. Generated code calls these delegates instead of referencing types directly.

## Hypothesis

Setting `Pack="false"` for `FormulaBoss.Runtime.dll` in the `.dna` file would give it a stable disk-backed identity. Roslyn's `MetadataReference.CreateFromFile()` would point to the same DLL that the runtime loaded, so the JIT could match type references.

## Test Method

Registered 8 diagnostic UDFs (FBTEST0–FBTEST7) from `AddIn.InitializeInterception()`, compiled through the same `DynamicCompiler` pipeline as real UDFs.

| UDF | What it tests | References host types? |
|-----|--------------|----------------------|
| FBTEST0 | Returns `"hello from generated code"` | No |
| FBTEST1 | Reports loaded assembly locations via `AppDomain.GetAssemblies()` | No (reflection only) |
| FBTEST2 | Calls `ExcelValue.Wrap()` via `Assembly.GetType()` + `MethodInfo.Invoke()` | No (reflection only) |
| FBTEST3 | Loads `ExcelReference` via `Assembly.GetType()` | No (reflection only) |
| FBTEST4 | `typeof(FormulaBoss.Runtime.ExcelValue)` | **Yes — direct type ref** |
| FBTEST5 | `typeof(ExcelDna.Integration.ExcelReference)` | **Yes — direct type ref** |
| FBTEST6 | `FormulaBoss.Runtime.ExcelValue.Wrap(42.0)` | **Yes — direct method call** |
| FBTEST7 | `ExcelValue.Wrap(arr)` + `.Rows.Count()` | **Yes — direct method call** |

Each test was in a separate class to isolate JIT failures.

## Configuration

`.dna` file during test:
```xml
<ExternalLibrary Path="formula-boss.dll" Pack="false" />
<Reference Path="ExcelDna.Integration.dll" Pack="false" />
<Reference Path="FormulaBoss.Runtime.dll" Pack="false" />
```

## Results

**FBTEST0:** `"hello from generated code"` — PASS

**FBTEST1 (assembly locations):**
```
Runtime: C:\Users\trjac\Repositories\taglo-formula-boss\formula-boss\bin\Debug\net6.0-windows\FormulaBoss.Runtime.dll
  FullName: FormulaBoss.Runtime, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
ExcelDNA: (empty location)
  FullName: ExcelDna.Integration, Version=1.1.0.0, Culture=neutral, PublicKeyToken=f225e9659857edbe
Host: C:\Users\trjac\Repositories\taglo-formula-boss\formula-boss\bin\Debug\net6.0-windows\formula-boss.dll
  FullName: formula-boss, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
```

Key observations:
- `FormulaBoss.Runtime` has a valid disk location (Pack="false" is working)
- `ExcelDna.Integration` has an empty location despite `Pack="false"` — ExcelDNA's loader handles it differently
- `DynamicCompiler.AddRuntimeReference()` uses `typeof(Runtime.ExcelValue).Assembly.Location` which points to the on-disk DLL

**FBTEST2:** `PASS: Wrap(42.0) returned FormulaBoss.Runtime.ExcelScalar` — reflection works perfectly

**FBTEST3:** `PASS: ExcelReference found` — reflection works for ExcelDNA types too

**FBTEST4–FBTEST7:** All compiled successfully (registered as UDFs) but returned `#VALUE!` at runtime.

## Conclusion

**`Pack="false"` does not solve the assembly identity mismatch for stream-loaded assemblies.**

Even when:
- The Runtime DLL is on disk with a valid `Location`
- Roslyn compiles against that exact file via `MetadataReference.CreateFromFile(assembly.Location)`
- The types, versions, and public key tokens all match

...the JIT still cannot resolve type references from an assembly loaded via `AssemblyLoadContext.Default.LoadFromStream()`. The mismatch is inherent to how `LoadFromStream()` creates assemblies without file-backed identity.

## Root Cause

`LoadFromStream()` creates an assembly in the default `AssemblyLoadContext` but without a file path association. When the JIT encounters a type reference like `FormulaBoss.Runtime.ExcelValue`, it needs to resolve the assembly `FormulaBoss.Runtime`. The resolution logic cannot match the in-memory stream-loaded assembly's reference metadata to the already-loaded copy of `FormulaBoss.Runtime.dll`, even though they are byte-identical. This is a .NET runtime limitation, not an ExcelDNA-specific issue.

## Implications

1. **The delegate bridge pattern is mandatory** for all cross-assembly interaction from generated code — not just ExcelDNA, but any host-loaded assembly including `FormulaBoss.Runtime`.
2. **Wrapper types are still viable** — they work at runtime in the host context. Generated code accesses them via delegate bridges (all `object` signatures) or reflection.
3. **Roslyn intellisense is not affected** — the editor's synthetic document is not stream-loaded; it uses Roslyn's workspace API which handles references differently.

## Attempted Fix That Did NOT Work

Adding `<Reference Path="FormulaBoss.Runtime.dll" Pack="false" />` to the `.dna` file. This correctly placed the DLL on disk and gave it a valid `Assembly.Location`, but did not fix the JIT type resolution.
