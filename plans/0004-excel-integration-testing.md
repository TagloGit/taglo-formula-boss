# 0004 — Excel Integration Testing — Implementation Plan

## Overview

Add a new `formula-boss.AddinTests` project that loads the Formula Boss XLL into a real Excel instance and tests the full pipeline end-to-end. We'll spike ExcelDna.Testing 1.9.0 first (it supports net6.0-windows and provides xUnit integration out of the box). If it doesn't work with our stack, fall back to custom COM automation via `RegisterXLL`.

## Phase 1 — ExcelDna.Testing Spike

### Files to Create

- `formula-boss.AddinTests/formula-boss.AddinTests.csproj` — New test project targeting `net6.0-windows`
- `formula-boss.AddinTests/SmokeTests.cs` — First end-to-end tests using ExcelDna.Testing

### Files to Modify

- `formula-boss/formula-boss.slnx` — Add the new project

### Project Setup

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net6.0-windows</TargetFramework>
    <RootNamespace>FormulaBoss.AddinTests</RootNamespace>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsTestProject>true</IsTestProject>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="ExcelDna.Testing" Version="1.9.0" />
  </ItemGroup>
</Project>
```

Key: ExcelDna.Testing brings in xUnit, the test SDK, and COM interop references automatically.

### Assembly Attribute

```csharp
[assembly: Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]
```

### Test Class Shape

```csharp
[ExcelTestSettings(AddIn = @"..\formula-boss\bin\Debug\net6.0-windows\formula-boss-AddIn")]
public class SmokeTests
{
    [ExcelFact]
    public void SimpleExpression_ReturnsCorrectValue()
    {
        var ws = Util.Workbook.Sheets[1];
        // Set up source data
        ws.Range["A1"].Value = 10;
        ws.Range["A2"].Value = 20;
        ws.Range["A3"].Value = 30;

        // Write a formula-boss formula that the interceptor will pick up
        // The interceptor watches for =FB. prefix formulas
        ws.Range["B1"].Formula2 = "=FB.sum(A1:A3)";  // or whatever triggers interception

        // Wait for interception + compilation
        // Poll until result appears or timeout
        // Assert result
    }
}
```

### Spike Deliverables

1. **Does `dotnet test` successfully launch Excel and load the XLL?** — If ExcelDna.Testing can't find or load our add-in, we know immediately.
2. **Can we trigger formula interception from test code?** — Write a `=FB.` formula via COM and see if the interceptor fires.
3. **What's the timing?** — Measure how long interception + Roslyn compilation takes, determine polling strategy.

### Spike Decision Gate

If any of these fail:
- ExcelDna.Testing can't load a net6.0 XLL
- In-process mode conflicts with our ALC assembly loading
- COM reference issues on net6.0-windows

Then proceed to **Phase 1b — Custom COM Fallback** (below).

## Phase 1b — Custom COM Fallback (only if spike fails)

Replace ExcelDna.Testing with manual COM automation:

```csharp
public sealed class ExcelAddinFixture : IAsyncLifetime
{
    private dynamic _app;

    public async Task InitializeAsync()
    {
        var type = Type.GetTypeFromProgID("Excel.Application");
        _app = Activator.CreateInstance(type);
        _app.Visible = false;
        _app.DisplayAlerts = false;

        // Load the XLL
        var xllPath = FindXllPath();
        _app.RegisterXLL(xllPath);

        // Wait for AutoOpen to complete
        await Task.Delay(2000);
    }

    public Task DisposeAsync() { /* Quit + ReleaseComObject */ }
}
```

## Phase 2 — Core Test Suite

Once the spike succeeds with either approach, build out tests:

### Files to Create/Modify

- `formula-boss.AddinTests/SmokeTests.cs` — Expand with Tier 1 tests
- `formula-boss.AddinTests/PipelineTests.cs` — Tier 2 core pipeline tests
- `formula-boss.AddinTests/TestUtilities.cs` — Shared helpers (polling, range setup, cleanup)

### Test Helpers Needed

```csharp
/// <summary>
/// Writes a formula and waits for the interceptor to compile and register the UDF.
/// Polls Range.Value until it's no longer #NAME? or timeout.
/// </summary>
static object? WriteFormulaAndWait(dynamic worksheet, string cell, string formula,
    int timeoutMs = 10000, int pollIntervalMs = 200)
```

### Tier 1 — Smoke Tests
1. **Scalar result:** `=FB.sum(A1:A3)` on `{10, 20, 30}` → `60`
2. **Array spill:** `=FB.data.toArray()` → spilled range matches input
3. **Interception fires:** Confirm formula cell doesn't show `#NAME?` after timeout

### Tier 2 — Core Pipeline Tests
4. **Object model path:** Set cell colors, write `=FB.data.cells.where(c => c.color == 6).sum()` → correct result
5. **Multi-input LET:** `=FB.LET(x, A1:A3, y, B1:B3, x + y)` (if applicable to current DSL)
6. **Table with headers:** Range with header row, `=FB.data.col("Price").sum()`
7. **Error expression:** Invalid DSL → cell shows error string (not crash)

### Tier 3 — Deferred (future issue)
- Re-edit formula updates UDF
- Multiple coexisting formulas
- Large range performance

## Order of Operations

1. **Create project + spike** — New csproj, assembly attribute, one smoke test. Run `dotnet test` and see what happens. Record findings in a spike notes section of the PR description.
2. **Decide approach** — ExcelDna.Testing works → continue. Doesn't work → implement Phase 1b.
3. **Add polling helper** — `WriteFormulaAndWait` utility for reliable async assertion.
4. **Build Tier 1 tests** — Three smoke tests that prove the loop works.
5. **Build Tier 2 tests** — Core pipeline coverage.
6. **Update solution** — Add project to `.slnx`, ensure `dotnet test` at solution level runs all tests (including existing ones).

## Testing Approach

- Tests in this project are inherently integration tests — they require Excel to be installed
- `dotnet test --filter "FullyQualifiedName~AddinTests"` to run only add-in tests
- `dotnet test --filter "FullyQualifiedName!~AddinTests"` to skip them (CI or no-Excel environments)
- Each test class should clean up its worksheet data (or use fresh sheets) to avoid interference

## XLL Path Discovery

The XLL path depends on the build output. Strategy:
1. Compute relative to test assembly location: `../formula-boss/bin/Debug/net6.0-windows/formula-boss-AddIn64.xll`
2. Or use `ExcelTestSettings(AddIn = ...)` with a relative path (ExcelDna.Testing resolves bitness automatically)
3. The project must build `formula-boss` first — add a `<ProjectReference>` or rely on solution build order

## Risks

| Risk | Mitigation |
|---|---|
| ExcelDna.Testing doesn't work with net6.0 | Phase 1b fallback ready |
| Interception timing is unpredictable | Polling with generous timeout (10s), configurable |
| Excel process leaks on test failure | `IDisposable`/`IAsyncLifetime` with process kill fallback |
| Tests interfere with user's open Excel | Launch a separate hidden instance; never attach to running Excel |
