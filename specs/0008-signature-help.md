# 0008 — Signature Help (Parameter Info)

## Problem

When users type method calls in DSL expressions (e.g. `table.Rows().Where(|`), there is no indication of what parameters the method expects. Users must memorize or look up the wrapper type APIs externally. This is a significant usability gap — VS-style signature help (parameter info) is one of the most impactful intellisense features for discoverability and reducing errors.

Additionally, the Runtime wrapper types (`ExcelValue`, `ExcelArray`, `RowCollection`, `Row`, `ColumnValue`, etc.) have essentially zero XML documentation. Even once signature help is wired up, there would be nothing meaningful to display beyond raw type signatures.

## Proposed Solution

Add VS-style signature help to the AvalonEdit editor: when the cursor is inside a method's argument list, display a tooltip-style popup showing the method signature with the active parameter highlighted. Support overload navigation for methods with multiple signatures.

As a prerequisite, add XML doc comments (`<summary>`, `<param>`, `<returns>`) to all user-facing public methods on the Runtime wrapper types so that signature help displays meaningful descriptions.

## User Stories

- As a Formula Boss user, I want to see method signatures when I type `(` or `,` inside a method call, so that I know what parameters are expected without leaving the editor.
- As a Formula Boss user, I want to navigate between overloads (e.g. `First()` vs `First(predicate)`) so that I can pick the right variant.
- As a Formula Boss user, I want to see parameter descriptions in the signature popup so that I understand what each argument does.

## Behaviour

### Trigger

The signature help popup appears when:
- The user types `(` after a method name inside a DSL backtick region
- The user types `,` inside an argument list
- The user positions the cursor inside an existing argument list (e.g. clicking or arrowing into it)

### Display

A lightweight popup appears above the current line (or below if insufficient space above), showing:

```
RowCollection.Where(Func<dynamic, bool> predicate)
                    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Filters rows to those matching the predicate.

predicate: A function that returns true for rows to keep.
```

- The active parameter is highlighted (bold or underline).
- If overloads exist, show "1 of N" with Up/Down arrow navigation (matching VS behaviour).
- The popup is read-only — no interaction beyond overload navigation and dismiss.

### Dismiss

The popup dismisses when:
- The user types `)` that closes the argument list
- The user presses Escape
- The cursor moves outside the argument list (e.g. backspacing past the `(`)
- The editor loses focus

### Active Parameter Tracking

As the user types or moves the cursor within the argument list, the highlighted parameter updates based on comma position. For nested calls (e.g. `Select(r => r.Where(|))`), the innermost method's signature is shown.

### Scope

Signature help is available inside DSL backtick regions only — the same context where Roslyn-powered completions already work. The existing `SyntheticDocumentBuilder` and `RoslynWorkspaceManager` infrastructure handles the mapping.

Non-DSL contexts (top-level LET bindings, table name entry) do not get signature help — there are no method calls in those contexts.

## Technical Approach

### Roslyn Integration

Roslyn provides `SignatureHelpService` (or the equivalent `GetItemsAsync` on `SignatureHelpProvider`) which, given a document and cursor position, returns:
- List of signature overloads (each with parameter names, types, and XML doc descriptions)
- The active signature (best match for current arguments)
- The active parameter index

This operates on the same `AdhocWorkspace` and synthetic documents already used for completions and diagnostics. The `SyntheticDocumentBuilder` already maps cursor positions between the editor and the synthetic C# document.

### UI

A custom WPF popup (not the AvalonEdit `CompletionWindow` — that's for list selection). A simple styled `Popup` or `ToolTip`-derived control containing:
- Method signature text with the active parameter highlighted
- Overload counter ("1 of 3") with Up/Down key handling
- Summary and parameter description text

Positioned relative to the caret, similar to how `ErrorHighlighter` positions its tooltips.

The signature help popup coexists with the completion window — both can be visible simultaneously (VS does this: completion list below the cursor, signature help above).

### XML Documentation

For Roslyn to surface parameter descriptions, the Runtime types need `<summary>`, `<param>`, and `<returns>` XML doc comments. The `.csproj` should enable `<GenerateDocumentationFile>true</GenerateDocumentationFile>` so Roslyn can read the XML at analysis time.

**Types requiring XML docs (user-facing public API only):**

| Type | Approximate method count |
|---|---|
| `ExcelValue` | ~25 (collection methods + aggregations) |
| `ExcelScalar` | Inherits from ExcelValue — may need overrides documented if signatures differ |
| `ExcelArray` | Inherits — same consideration |
| `ExcelTable` | `Headers` property |
| `RowCollection` | ~17 methods |
| `GroupedRowCollection` | ~9 methods |
| `Row` | Indexers, `ColumnCount` |
| `ColumnValue` | `Value`, `Cell`, implicit conversions |

Operators (`+`, `-`, `==`, etc.) and standard overrides (`ToString`, `Equals`, `GetHashCode`) do not need XML docs — they don't appear in signature help contexts.

## Acceptance Criteria

- [ ] Typing `(` after a method name inside a DSL expression shows a signature help popup with the method's parameter list
- [ ] Typing `,` updates the active (highlighted) parameter
- [ ] Overloaded methods show "N of M" with Up/Down arrow navigation between overloads
- [ ] Signature popup displays `<summary>` text for the method and `<param>` text for the active parameter
- [ ] Popup dismisses on `)`, Escape, cursor leaving the argument list, or editor losing focus
- [ ] Popup coexists with the completion window (both can be visible)
- [ ] All user-facing public methods on Runtime wrapper types have XML doc comments
- [ ] `formula-boss.Runtime.csproj` generates an XML documentation file

## Out of Scope

- Signature help outside DSL backtick regions (no method calls in those contexts)
- Signature help for operators (e.g. `+`, `==`)
- Inline parameter name hints (ghost text showing parameter names at call sites — separate feature)
- Quick Info on hover (showing type/doc info for any symbol under the cursor — separate feature)

## Open Questions

None — ready for review.
