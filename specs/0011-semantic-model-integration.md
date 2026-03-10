# 0011 — Semantic Model Integration

## Status

Draft

## Problem

The transpilation pipeline makes type-dependent decisions using **syntax pattern matching** — looking for specific AST shapes (e.g. `r["Col"]` inside a lambda) to infer that a parameter is a table and needs `[#All]` header inclusion. This approach breaks on any usage pattern that doesn't match the expected shape:

- `tbl.Rows.First(r => ...)["Distance"]` — bracket access outside a lambda
- `var row = tbl.Rows.First(); row["Col"]` — bracket access on a local variable
- Passing a table to a helper that does the bracket access internally

Each new C# pattern the user writes risks becoming another edge case.

Meanwhile, the editor (`SyntheticDocumentBuilder` + `RoslynWorkspaceManager`) **already maintains a typed compilation context** with table types, row types, and column properties. It already produces Roslyn semantic models for diagnostics and signature help. But this rich type information isn't used by the pipeline and isn't surfaced to the user via hover.

## Goals

1. **Robust header detection** — determine which parameters need `[#All]` using type information rather than AST pattern matching.
2. **Hover type display** — show resolved types when the user hovers over `var` declarations or identifiers in the editor.
3. **Shared infrastructure** — unify the type analysis used by the pipeline and editor so both benefit from the same semantic understanding.

## Non-goals

- Full IntelliSense rewrite (the existing completion system works well).
- Changing the runtime type system (`ExcelValue`, `Row`, `ColumnValue`).
- Type inference for parameters that aren't tables or named ranges (scalar parameters stay as `ExcelValue`).

## Design

### 1. Metadata-aware header detection in the pipeline

**Current state:** `InputDetector.DetectHeaderVariables()` scans the AST for `ElementAccessExpressionSyntax` with string literals and traces them back through lambda parameters. `FormulaPipeline` then appends `[#All]` to any parameter in the header variables set.

**Problem:** The detection is fragile because it pattern-matches syntax rather than understanding types.

**Proposed change:** `FormulaPipeline` already has access to workbook metadata through `FormulaInterceptor` (which holds a `WorkbookMetadataProvider`). Pass the metadata into `FormulaPipeline.Process()` via `ExpressionContext` and use it to determine which parameters are tables:

```
ExpressionContext
├── PreferredUdfName: string?     (existing)
└── Metadata: WorkbookMetadata?   (new)
```

The header variable decision becomes:

```csharp
// A parameter needs [#All] if:
// 1. Its name matches a known Excel table name in the workbook, OR
// 2. It was detected by the existing AST pattern matching (keep as fallback)
var headerVariables = detection.Parameters
    .Where(p => metadata?.IsTable(p) == true || detection.HeaderVariables.Contains(p))
    .ToHashSet();
```

This handles the common case (direct table references) without needing a semantic model in the pipeline. The existing AST detection stays as a fallback for edge cases where the parameter name doesn't match a table name (e.g. a LET-aliased table).

**Impact on `InputDetector`:** `DetectHeaderVariables()` remains but is no longer the sole source of truth. Over time, as the semantic model matures, it could be removed entirely.

### 2. Semantic model service (shared between pipeline and editor)

Extract the core compilation logic from `RoslynWorkspaceManager` and `SyntheticDocumentBuilder` into a shared `SemanticAnalysisService`:

```
SemanticAnalysisService
├── BuildTypedCompilation(expression, metadata) → SemanticModel
├── GetTypeAtPosition(semanticModel, offset) → ITypeSymbol?
├── GetExpressionType(semanticModel, expression) → ITypeSymbol?
└── IsTableType(semanticModel, parameterName) → bool
```

**Key design decisions:**

- **Reuse `SyntheticDocumentBuilder`'s typed classes** — the `__tblCastlesTable : ExcelTable` pattern already works for completions and diagnostics. The same synthetic types work for semantic queries.
- **Don't create a persistent workspace per expression** — create a one-shot `CSharpCompilation` for each analysis request. The pipeline processes one expression at a time; workspace management overhead isn't justified.
- **Cache the compilation** — the pipeline already caches UDF results by expression. Cache the semantic model alongside it for editor queries on the same expression.

**Where this lives:**

```
formula-boss/
  Analysis/
    SemanticAnalysisService.cs    — shared service
    TypedDocumentBuilder.cs       — extracted from SyntheticDocumentBuilder
  UI/Completion/
    SyntheticDocumentBuilder.cs   — delegates to TypedDocumentBuilder
    RoslynWorkspaceManager.cs     — uses SemanticAnalysisService for type queries
  Transpilation/
    InputDetector.cs              — unchanged (fallback detection)
  Interception/
    FormulaPipeline.cs            — uses SemanticAnalysisService for header detection
```

### 3. Hover type display in the editor

**Current state:** `ErrorHighlighter` already handles mouse hover — it shows red squiggle tooltips with Roslyn diagnostic messages. The infrastructure for position-based tooltips exists.

**Proposed change:** Extend hover handling to show type information when the cursor is over:

- A `var` keyword — show the inferred type (e.g. `IExcelRange`, `IEnumerable<ColumnValue>`)
- An identifier — show its type (e.g. `tblCastles: ExcelTable`, `country: ExcelValue`)
- A lambda parameter — show its type (e.g. `r: __tblCastlesRow` displayed as `Row {Country, Castle, ...}`)

**Implementation:**

1. On mouse hover, map editor offset → synthetic document offset (reuse `DiagnosticBuildResult` position mapping).
2. Find the `SyntaxNode` at that offset.
3. Query `semanticModel.GetTypeInfo(node)` or `semanticModel.GetSymbolInfo(node)`.
4. Format the type name for display:
   - Strip synthetic prefixes (`__tblCastlesRow` → `Row`)
   - For row types, show column names: `Row {Country, Castle, Distance}`
   - For collection types, show element type: `IEnumerable<Row {Country, Castle, ...}>`
   - For `ExcelValue` subtypes, show the specific subtype
5. Display in a tooltip near the cursor (reuse `ErrorHighlighter`'s tooltip infrastructure).

**Type name formatting rules:**

| Synthetic type | Displayed as |
|---|---|
| `__tblCastlesTable` | `ExcelTable (tblCastles)` |
| `__tblCastlesRow` | `Row {Country, Castle, Distance, ...}` |
| `__tblCastlesRowCollection` | `RowCollection (tblCastles)` |
| `ColumnValue` | `ColumnValue` |
| `ExcelValue` | `ExcelValue` |
| `ExcelScalar` | `ExcelScalar` |
| `IExcelRange` | `IExcelRange` |

### 4. Pipeline integration — semantic model for header detection (future)

Once the shared `SemanticAnalysisService` exists, the pipeline can optionally use it for more sophisticated header detection:

```csharp
// Instead of just checking table names:
var semanticModel = analysisService.BuildTypedCompilation(expression, metadata);
var headerVariables = detection.Parameters
    .Where(p => analysisService.IsTableType(semanticModel, p))
    .ToHashSet();
```

This handles cases the metadata approach misses:
- LET-aliased tables: `=LET(subset, tblCastles, \`subset.Rows...\`)`  — `subset` isn't a table name, but the semantic model knows its type flows from `tblCastles`.
- Chained expressions: when one expression's result feeds into another.

This is deferred to a later phase since the metadata approach covers the immediate need.

## Rollout

### Phase 1 — Metadata-aware header detection

- Extend `ExpressionContext` with `WorkbookMetadata?`
- Pass metadata from `FormulaInterceptor` through to `FormulaPipeline`
- Use metadata table names alongside existing AST detection for `[#All]` decisions
- **Addresses the immediate bug** — tables like `tblCastDist` get `[#All]` even when bracket access is outside a lambda

### Phase 2 — Hover type display

- Extract shared `TypedDocumentBuilder` from `SyntheticDocumentBuilder`
- Build `SemanticAnalysisService` using one-shot compilations
- Add hover handler to `FloatingEditorWindow` / `ErrorHighlighter`
- Format and display type tooltips

### Phase 3 — Semantic header detection (optional)

- Use `SemanticAnalysisService` in `FormulaPipeline` for type-aware header detection
- Handles LET aliases and chained expressions
- Can eventually replace `DetectHeaderVariables()` entirely

## Decisions

1. **No shared compilation between pipeline and editor** — the pipeline runs a one-shot compilation on formula confirmation; the editor maintains a persistent `AdhocWorkspace` for continuous diagnostics. Different lifecycles, negligible performance gain from sharing. Keep them independent.

2. **Don't show column types in hover** — show column names only (e.g. `Row {Country, Castle, Distance}`). If the user accesses a column, they can hover over that access to see the column's type.

3. **Lambda parameter hover shows type name only** — for `r` in `tbl.Rows.Where(r => ...)`, show `Row (tblCastles)` rather than listing all columns.
