# Excel Inline UDF Add-in — Technical Stack and Implementation Plan

## Recommended Tech Stack

### Core Framework: ExcelDNA

**Why ExcelDNA:**

- Native C API integration — fastest possible UDF execution
- First-class support for dynamic UDF registration at runtime
- `AllowReference` attribute enables receiving range references (required for object model access)
- Mature, well-documented, actively maintained
- Single `.xll` file deployment — no installer complexity
- Supports custom task panes and ribbon UI

**Alternative considered:** VSTO

- Rejected because UDFs go through COM interop (slower)
- VSTO designed for automation and UI, not high-performance worksheet functions

---

### Runtime Compilation: Roslyn

**Why Roslyn (Microsoft.CodeAnalysis.CSharp):**

- Official C# compiler as a library
- In-memory compilation to assembly — no temp files
- Fast incremental compilation after initial load
- Rich error diagnostics for user feedback
- Well-documented API

**Usage pattern:**

- Generate C# source string from transpiled DSL
- Compile to in-memory assembly
- Extract method via reflection
- Register with ExcelDNA's dynamic registration API

---

### Excel Object Model Access: Microsoft.Office.Interop.Excel

**Why COM Interop:**

- Required for accessing cell properties not exposed through C API (colour, formatting, etc.)
- ExcelDNA provides `ExcelDnaUtil.Application` to get the `Application` object
- Standard approach, well-understood performance characteristics

**Considerations:**

- Object model calls are slow — design DSL to minimise when possible
- Auto-detect whether expression needs object model or can use fast value path

---

### Floating Editor (Post-MVP): AvalonEdit

**Why AvalonEdit:**

- WPF-native — integrates cleanly with ExcelDNA's task pane infrastructure
- Syntax highlighting with custom grammar definitions
- Lightweight and fast-loading
- Supports folding, bracket matching, basic autocomplete
- MIT licensed

**Alternative considered:** Monaco (VS Code editor)

- Richer features, true intellisense capability
- Rejected for MVP because requires WebView2 hosting, heavier weight, more complex integration
- Could revisit for future "deluxe" version

---

### DSL Parser: Hand-written Recursive Descent

**Why hand-written:**

- DSL is intentionally small and simple
- Recursive descent is easy to understand, debug, extend
- No external dependencies
- Full control over error messages

**Alternative considered:** Parser generator (ANTLR, etc.)

- Overkill for a small grammar
- Adds build complexity and dependencies
- Harder to produce friendly error messages

---

### Persistence (Post-MVP): JSON + File System

**Why JSON:**

- Human-readable for debugging
- Easy to edit manually if needed
- Standard serialisation, no special libraries

**Storage location:**

- User's AppData folder
- One file per named UDF or single library file

---

## Implementation Plan

### Phase 0: Project Setup

**Deliverables:**

- ExcelDNA project scaffolding
- Build pipeline producing `.xll` file
- Basic ribbon button that shows "Hello World" message box
- Verify add-in loads in Excel without errors

**Estimated effort:** Half day

---

### Phase 1: Static UDF Proof of Concept

**Deliverables:**

- Hand-written `FilterByColor(range, colorIndex)` UDF
- Demonstrates receiving range reference via `AllowReference`
- Demonstrates object model access to read cell colours
- Demonstrates returning 2D array result
- Verify performance with 1000+ cells

**Purpose:** Validate core technical approach before building transpiler.

**Estimated effort:** 1 day

---

### Phase 2: Roslyn Integration

**Deliverables:**

- Roslyn compilation pipeline
- Generate UDF from hardcoded C# source string
- Register dynamically with ExcelDNA
- Call generated UDF from Excel formula
- Measure and document compilation latency

**Purpose:** Validate dynamic compilation and registration works.

**Estimated effort:** 1-2 days

---

### Phase 3: DSL Parser — Core

**Deliverables:**

- Tokeniser for DSL expressions
- Parser for core constructs:
  - `range.cells`, `range.values`
  - `.where(predicate)`
  - `.select(transform)`
  - `.toArray()`
  - Cell properties: `.value`, `.color`, `.row`, `.col`
  - Lambda syntax: `c => expression`
  - Basic operators and literals
- AST representation

**Estimated effort:** 2-3 days

---

### Phase 4: Transpiler — Core

**Deliverables:**

- AST to C# source code generator
- Template for UDF method structure
- Handle range reference input
- Handle object model vs. value-only detection
- Generate correct LINQ chains
- Handle output normalisation to 2D array

**Estimated effort:** 2-3 days

---

### Phase 5: Formula Interception — Quote Prefix Path

**Deliverables:**

- Worksheet change event handler
- Detect cells starting with `'=` containing backticks
- Extract backtick expressions
- Integrate parser and transpiler
- Rewrite cell formula with generated UDF
- Basic error handling: set cell to error value

**Estimated effort:** 1-2 days

---

### Phase 6: Error Feedback — Basic

**Deliverables:**

- Parse error messages with position information
- Compile error capture and formatting
- Runtime error capture
- Display via cell comment (simplest approach)
- Error type differentiation in cell value

**Estimated effort:** 1 day

---

### Phase 7: MVP Testing and Hardening

**Deliverables:**

- Test suite covering DSL constructs
- Edge case handling (empty ranges, single cells, errors in data)
- Performance benchmarking
- Documentation for MVP feature set

**Estimated effort:** 2-3 days

---

### Phase 8: Floating Editor — Basic

**Deliverables:**

- WPF window hosting AvalonEdit
- Keyboard shortcut registration (when not in edit mode)
- Pre-load existing cell formula
- Apply button triggers parse/compile/rewrite
- Cancel button closes without changes
- Basic syntax highlighting for keywords

**Estimated effort:** 2-3 days

---

### Phase 9: Floating Editor — Enhanced

**Deliverables:**

- Real-time error underlining
- Autocomplete for DSL keywords and methods
- Signature help tooltips
- Bracket matching
- Expression history (session-scoped)

**Estimated effort:** 3-5 days

---

### Phase 10: Named UDF Registration

**Deliverables:**

- Dialog to name a UDF from current cell's expression
- Session-scoped name registry
- Name collision detection
- Use named UDF in subsequent formulas

**Estimated effort:** 1-2 days

---

### Phase 11: UDF Persistence

**Deliverables:**

- Save named UDF to JSON library file
- Load library on add-in startup
- Recompile or load cached assemblies
- UI to manage library (view, delete, rename)

**Estimated effort:** 2-3 days

---

### Phase 12: VBA Transpiler Backend

**Deliverables:**

- VBA code generator from DSL AST (separate from C# generator)
- LINQ-style operations converted to VBA loops
- Object model access translated to native VBA syntax
- VBA injection into workbook via `VBProject.VBComponents`
- "Prepare for Export" ribbon button
- Export dialog with options (transpile to VBA, bake to values, both)
- Documentation of Trust Center requirements

**Technical notes:**

- Reuses existing parser and AST from Phase 3
- New backend visitor/generator for VBA output
- Function names must match ExcelDNA-registered names
- Handle edge cases: empty results, error propagation

**Estimated effort:** 3-4 days

---

### Phase 13: Built-in Algorithm Library

**Deliverables:**

- Graph algorithms: shortest path, connected components, topological sort
- Combinatorics: permutations, combinations, subsets
- Iteration: recursive until condition
- Expose via DSL syntax
- VBA equivalents for all algorithms (for export)
- Documentation for each algorithm

**Estimated effort:** 3-5 days

---

### Phase 14: Polish and Distribution

**Deliverables:**

- Installer or single-file distribution
- User documentation / quick reference card
- Performance tuning
- Error message refinement
- Versioning and update mechanism

**Estimated effort:** 2-3 days

---

## Milestones Summary

| Milestone | Phases | Description | Estimated Total |
|-----------|--------|-------------|-----------------|
| **M1: Technical Validation** | 0-2 | Prove core tech works | 2-4 days |
| **M2: MVP** | 3-7 | Quote-prefix path, basic DSL, error feedback | 8-12 days |
| **M3: Editor** | 8-9 | Floating editor with highlighting and autocomplete | 5-8 days |
| **M4: Persistence** | 10-11 | Named UDFs, saved library | 3-5 days |
| **M5: Export** | 12 | VBA transpiler, portable workbooks | 3-4 days |
| **M6: Algorithms** | 13 | Built-in algorithm library | 3-5 days |
| **M7: Release** | 14 | Polish and distribution | 2-3 days |

**Total estimated effort to MVP:** 2-3 weeks of focused development

**Total estimated effort to full vision:** 7-9 weeks

---

## Risk Register

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| ExcelDNA dynamic registration doesn't work as expected | Low | High | Phase 2 validates early |
| Roslyn compilation too slow for interactive use | Low | Medium | Caching, lazy load, measure in Phase 2 |
| Object model access slower than acceptable | Medium | Medium | Document limits, optimise hot paths, prefer value path |
| DSL grammar ambiguities discovered late | Medium | Medium | Keep grammar simple, test thoroughly in Phase 3 |
| Keyboard shortcut conflicts with Excel or other add-ins | Medium | Low | Make shortcut configurable |
| AvalonEdit integration issues with Excel | Low | Medium | Fallback to simple WPF TextBox if needed |
| Formula bar edit-mode interception truly impossible | High | Low | Already accepted — quote-prefix workaround |
| VBA project access blocked by Trust settings | Medium | Medium | Document requirement, detect and prompt user |
| VBA transpilation produces incorrect results | Medium | High | Comprehensive test suite comparing C# and VBA outputs |

---

## Dependencies

**Required:**

- Visual Studio 2022 or later (or Rider)
- .NET Framework 4.7.2+ or .NET 6+ (ExcelDNA supports both)
- ExcelDNA NuGet package
- Microsoft.CodeAnalysis.CSharp NuGet package
- Microsoft.Office.Interop.Excel (installed with Office)

**Optional:**

- AvalonEdit NuGet package (for floating editor)
- Newtonsoft.Json or System.Text.Json (for persistence)

---

## Open Technical Decisions

1. **Target framework:** .NET Framework 4.7.2 (maximum compatibility) or .NET 6+ (modern, but requires recent ExcelDNA)?

2. **UDF naming scheme:** Hash of expression, incrementing counter, or user-facing descriptive names?

3. **Caching strategy:** Cache compiled assemblies to disk, or always recompile on session start?

4. **Task pane vs. floating window:** Task pane docks in Excel (more integrated), floating window can be positioned freely (more flexible). Which for editor?

5. **Shortcut key:** What's an ergonomic, unoccupied shortcut? `Ctrl+Shift+E`? `Ctrl+\``? Configurable from start?

---

## Next Steps

1. Create ExcelDNA project, verify builds and loads
2. Implement Phase 1 proof-of-concept UDF
3. Validate Phase 2 Roslyn integration
4. Review and refine DSL grammar before Phase 3
5. Proceed to MVP phases
