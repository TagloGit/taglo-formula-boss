# 0004 — Excel Integration Testing

## Problem

Claude agents can run unit tests and compile code, but cannot verify that Formula Boss actually works end-to-end inside Excel. The existing `formula-boss.IntegrationTests` project tests DSL → compile → execute against real Excel COM ranges, but it bypasses the add-in entirely — it never loads the XLL, never tests formula interception, UDF registration, or dynamic array spilling.

This means every feature that touches the Excel integration boundary (new DSL functions, interception changes, spill behavior, macro-type registration) requires Tim to manually open Excel and test. Closing this loop would let Claude agents ship higher-confidence PRs and catch regressions that unit tests miss.

## Proposed Solution

Add a new test project (`formula-boss.AddinTests`) that loads the Formula Boss XLL add-in into a real Excel instance and tests the full pipeline: type a formula → interceptor fires → UDF compiles and registers → cell shows correct result.

### Approach: ExcelDna.Testing First, Custom COM Fallback

**Phase 1 — Spike ExcelDna.Testing 1.9.0.** This is a purpose-built xUnit framework that launches Excel, loads an XLL, and provides COM access for assertions. If it works with our .NET 6 / ExcelDNA 1.9 stack, it eliminates significant boilerplate.

**Phase 2 — If ExcelDna.Testing doesn't work** (e.g., .NET 6 COM reference issues, in-process mode conflicts with our ALC loading), fall back to custom COM automation that launches Excel, loads the XLL via `Application.RegisterXLL()`, and drives formulas via the COM object model.

### Coexistence with Existing Tests

The existing `formula-boss.IntegrationTests` project stays as-is. It tests a different layer (compiled _Core methods against COM ranges, without the add-in). The new project tests the add-in-loaded layer. Both run via `dotnet test`.

## User Stories

- As a Claude agent, I want to run `dotnet test` and verify that a Formula Boss expression produces the correct result inside Excel, so I can validate my changes without asking Tim to test manually.
- As Tim, I want integration test failures to surface in PR review, so I can trust that Claude-authored PRs work end-to-end.
- As a developer, I want the test suite to handle Excel lifecycle automatically (launch, load add-in, clean up), so tests are self-contained and repeatable.

## Acceptance Criteria

- [ ] A new test project (`formula-boss.AddinTests`) exists in the solution
- [ ] Tests run via `dotnet test` using the standard xUnit runner
- [ ] The test framework launches Excel, loads the Formula Boss XLL, and shuts down Excel on completion
- [ ] At least one test verifies the full loop: write a `=FB.` formula → interceptor registers UDF → cell returns correct value
- [ ] At least one test verifies dynamic array spilling (multi-cell result)
- [ ] At least one test verifies the object model path (`.cells` with color filtering)
- [ ] Tests clean up Excel processes reliably, even on failure
- [ ] Existing `formula-boss.IntegrationTests` continue to pass unchanged
- [ ] A spike document records which approach was chosen (ExcelDna.Testing vs custom COM) and why

## Test Categories

### Tier 1 — Smoke Tests (spike deliverable)
- Formula interception triggers on cell edit
- Simple value-path expression returns correct scalar
- Simple value-path expression returns correct spilled array

### Tier 2 — Core Pipeline
- Object model path (`.cells.where(c => c.color == 6).sum()`) with pre-colored cells
- Multi-input LET formulas with multiple range references
- Table expressions with header detection
- Error expressions return `#VALUE!` or error string

### Tier 3 — Edge Cases (future)
- Re-editing a formula updates the UDF registration
- Multiple formulas in different cells coexist
- Large range performance (1000+ cells)

## Out of Scope

- CI/CD automation (requires Excel installed — local-only for now)
- Testing the floating editor UI (WPF, not automatable via COM)
- Performance benchmarking (separate concern)
- Testing on multiple Excel versions

## Technical Considerations

### ExcelDna.Testing API
- `[assembly: TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]` hooks the xUnit runner
- `[ExcelTestSettings(AddIn = @"path\to\formula-boss-AddIn")]` loads the XLL
- `Util.Application` / `Util.Workbook` provide COM access
- In-process mode runs tests inside Excel (C API access); out-of-process uses COM interop
- Need to evaluate: does it work with .NET 6 targets? Does in-process conflict with our ALC assembly loading?

### Custom COM Fallback
- `Type.GetTypeFromProgID("Excel.Application")` to launch Excel
- `app.RegisterXLL(xllPath)` to load the add-in
- Write formulas via `Range.Formula2`, read results via `Range.Value`
- Must handle: Excel process cleanup, XLL path discovery, build-before-test dependency

### Formula Interception Timing
- The interceptor fires on `SheetChange` events, then uses `QueueAsMacro` for compilation
- Tests will need to wait for async compilation to complete before asserting results
- Polling `Range.Value` with a timeout is the likely approach

## Open Questions

- Does ExcelDna.Testing work with .NET 6 target frameworks, or only .NET Framework? (The NuGet page mentions a .NET Framework COM reference issue — spike will answer this)
- What's the minimum wait time for formula interception + compilation to complete? Need to determine a reasonable polling timeout.
- Should the XLL path be auto-discovered from the build output, or configured via a test settings file?
