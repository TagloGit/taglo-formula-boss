# 0008 — Signature Help — Implementation Plan

## Overview

Add VS-style signature help (parameter info) to the floating editor. When the cursor is inside a method call's argument list within a DSL backtick region, a popup shows the method signature with the active parameter highlighted and XML doc descriptions. Overloaded methods show "N of M" with Up/Down navigation.

This requires two workstreams: (1) adding XML doc comments to all user-facing Runtime type methods, and (2) building the signature help UI and Roslyn integration.

## Workstream 1: XML Documentation

### 1a. Enable XML doc generation

**File:** `formula-boss.Runtime/formula-boss.Runtime.csproj`

Add `<GenerateDocumentationFile>true</GenerateDocumentationFile>` to the `<PropertyGroup>`. This produces `formula-boss.Runtime.xml` alongside the DLL, which Roslyn reads automatically when loading metadata references.

### 1b. Add XML docs to Runtime wrapper types

Add `<summary>`, `<param>`, and `<returns>` tags to all user-facing public methods and properties. Operators and standard overrides (`ToString`, `Equals`, `GetHashCode`) are excluded — they don't appear in signature help contexts.

**Files to touch (all in `formula-boss.Runtime/`):**

| File | Methods to document | Notes |
|---|---|---|
| `ExcelValue.cs` | ~22 abstract methods/properties | Base definitions — these docs apply everywhere |
| `ExcelScalar.cs` | 0 | Overrides only; inherits docs from ExcelValue |
| `ExcelArray.cs` | 2 (`RowCount`, `ColCount`) | Only the new properties; overrides inherit docs |
| `ExcelTable.cs` | 1 (`Headers`) | Only the new property |
| `RowCollection.cs` | ~16 methods | Separate docs from ExcelValue (different parameter types — `Func<dynamic, bool>` vs `Func<ExcelValue, bool>`) |
| `GroupedRowCollection.cs` | ~7 methods | Similar to RowCollection but for groups |
| `RowGroup.cs` | 1 (`Key`) | Only the new property |
| `Row.cs` | ~4 (indexers, `ColumnCount`) | Skip DynamicObject overrides |
| `ColumnValue.cs` | ~3 (`Value`, `Cell`, implicit conversions) | |

**Doc style — keep descriptions short and practical:**
```csharp
/// <summary>Filters elements to those matching the predicate.</summary>
/// <param name="predicate">A function that returns true for elements to keep.</param>
/// <returns>A new range containing only matching elements.</returns>
public abstract IExcelRange Where(Func<ExcelValue, bool> predicate);
```

**Note on `inheritdoc`:** C# supports `<inheritdoc/>` for override methods. However, Roslyn's `SignatureHelpService` resolves XML docs from the metadata XML file, and `<inheritdoc/>` is expanded by the compiler only when `GenerateDocumentationFile` is set. The abstract methods on `ExcelValue` define the canonical docs; overrides in `ExcelScalar`/`ExcelArray` can use `<inheritdoc/>` or omit docs entirely (Roslyn walks the hierarchy). Either way, no need to duplicate docs on every override.

## Workstream 2: Signature Help Integration

### 2a. Add `GetSignatureHelpAsync` to `RoslynWorkspaceManager`

**File:** `formula-boss/UI/Completion/RoslynWorkspaceManager.cs`

Add a new method parallel to `GetCompletionsAsync`:

```csharp
public async Task<SignatureHelpItems?> GetSignatureHelpAsync(
    string syntheticSource, int caretOffset, CancellationToken cancellationToken)
{
    UpdateDocument(syntheticSource);
    var document = _workspace.CurrentSolution.GetDocument(_documentId)!;
    var service = SignatureHelpService.GetService(document);
    if (service == null) return null;

    var triggerInfo = new SignatureHelpTriggerInfo(
        SignatureHelpTriggerReason.InvokeSignatureHelpCommand);
    return await service.GetItemsAsync(
        document, caretOffset, triggerInfo, cancellationToken: cancellationToken);
}
```

`SignatureHelpItems` (from `Microsoft.CodeAnalysis.SignatureHelp`) contains:
- `Items` — list of overloads, each with parameter names/types/docs and method summary
- `SelectedItemIndex` — best-match overload
- `ArgumentIndex` — which parameter the cursor is on

### 2b. Create `SignatureHelpProvider`

**New file:** `formula-boss/UI/Completion/SignatureHelpProvider.cs`

Orchestrates signature help requests, analogous to `RoslynCompletionProvider` for completions:

1. Check context — only proceed if inside DSL backticks (reuse `ContextResolver.IsInsideBackticks`)
2. Build synthetic document via `SyntheticDocumentBuilder.Build()` (same as completions)
3. Call `RoslynWorkspaceManager.GetSignatureHelpAsync()`
4. Map `SignatureHelpItems` to a view model (`SignatureHelpModel`) suitable for the UI

**`SignatureHelpModel` (in same file or separate):**
```csharp
public record SignatureHelpModel(
    IReadOnlyList<SignatureOverload> Overloads,
    int ActiveOverloadIndex,
    int ActiveParameterIndex);

public record SignatureOverload(
    string MethodName,
    string Summary,
    IReadOnlyList<ParameterInfo> Parameters,
    string? ReturnType);

public record ParameterInfo(
    string Name,
    string Type,
    string? Description);
```

### 2c. Create `SignatureHelpPopup` UI

**New file:** `formula-boss/UI/SignatureHelpPopup.cs`

A WPF `Popup` control (not an AvalonEdit `CompletionWindow`) that displays signature info near the caret.

**Layout (built in code, no XAML needed):**
```
┌──────────────────────────────────────────────┐
│ ▲ 1 of 3 ▼  Where(predicate): RowCollection │
│                                              │
│ Filters rows to those matching the predicate.│
│                                              │
│ predicate: A function that returns true for   │
│ rows to keep.                                │
└──────────────────────────────────────────────┘
```

- **Overload counter** ("1 of 3") shown only when multiple overloads exist.
- **Signature line:** Method name, parameter list with active parameter in **bold**, return type.
- **Description area:** Method summary, then active parameter description.
- **Positioning:** Use `AvalonEdit.TextArea.TextView.GetVisualPosition()` to place the popup above the current line. If insufficient space above, place below.
- **Styling:** Match the completion window aesthetic (same background, border, font).

**Key behaviours:**
- `Update(SignatureHelpModel)` — refreshes content (called on each keystroke within the arg list)
- `Show()` / `Hide()` — visibility control
- Up/Down arrow handling for overload cycling (when popup is visible, intercept these keys)

### 2d. Wire triggers and lifecycle in `EditorBehaviorHandler` and `FloatingEditorWindow`

**File:** `formula-boss/UI/EditorBehaviorHandler.cs`

Add callbacks:
- `SignatureHelpRequested` — fired when `(` is typed or cursor enters an argument list
- `SignatureHelpDismissRequested` — fired when `)` closes the arg list, Escape, or cursor leaves

In `OnTextEntered`:
- After inserting `(`: fire `SignatureHelpRequested`
- After inserting `,`: fire `SignatureHelpRequested` (to update active parameter)
- After inserting `)`: fire `SignatureHelpDismissRequested`

In `OnPreviewKeyDown`:
- Up/Down when signature help is visible: cycle overloads (consume the key event)
- Escape when signature help is visible: dismiss (consume the key event)

**File:** `formula-boss/UI/FloatingEditorWindow.xaml.cs`

Add `ShowSignatureHelp()` async method (parallel to `ShowCompletion()`):
1. Ensure workspace is initialized
2. Call `SignatureHelpProvider.GetSignatureHelpAsync()`
3. If result is non-null, show/update `SignatureHelpPopup`
4. If null, hide popup

Add `_signatureHelpPopup` field and lifecycle management:
- Created lazily on first use
- Hidden when editor closes or formula is applied
- Coexists with completion window (both can be visible)

**Re-trigger on caret movement:** Subscribe to `TextArea.Caret.PositionChanged`. When the caret moves inside an argument list (detected by checking if the character at the caret's context is between `(` and matching `)`), re-query signature help to update active parameter. When the caret moves outside, dismiss.

### 2e. Ensure completion and signature help coexist

Both popups can be visible simultaneously. Key interactions:
- Completion window appears below the caret; signature help appears above.
- Up/Down arrows: if completion window is open, they navigate completions (existing behaviour). If only signature help is open, they cycle overloads.
- Escape: closes the topmost popup (completion first, then signature help).
- Typing `)` closes both.

## Order of Operations

### PR 1: XML documentation on Runtime types

1. Add `<GenerateDocumentationFile>true</GenerateDocumentationFile>` to `formula-boss.Runtime.csproj`
2. Add XML docs to `ExcelValue.cs` (abstract methods — canonical docs)
3. Add XML docs to `ExcelArray.cs` (`RowCount`, `ColCount`)
4. Add XML docs to `ExcelTable.cs` (`Headers`)
5. Add XML docs to `RowCollection.cs`
6. Add XML docs to `GroupedRowCollection.cs`
7. Add XML docs to `RowGroup.cs` (`Key`)
8. Add XML docs to `Row.cs`
9. Add XML docs to `ColumnValue.cs`
10. Build and verify no warnings from malformed XML

**Why separate PR:** XML docs are independently valuable (they'll also improve completion item descriptions if Roslyn surfaces them). No UI changes, low risk.

### PR 2: Roslyn signature help integration

1. Add `GetSignatureHelpAsync` to `RoslynWorkspaceManager`
2. Create `SignatureHelpProvider` with `SignatureHelpModel` types
3. Unit test: given a synthetic document with cursor inside `Where(|`, verify Roslyn returns the expected signature with parameter info

### PR 3: Signature help popup UI and editor wiring

1. Create `SignatureHelpPopup` WPF control
2. Add trigger/dismiss callbacks to `EditorBehaviorHandler`
3. Wire `ShowSignatureHelp()` in `FloatingEditorWindow`
4. Handle overload navigation (Up/Down)
5. Handle coexistence with completion window
6. AddIn test: open editor, type a method call with `(`, verify popup appears with correct signature

## Testing Approach

**Unit tests (`formula-boss.Tests`):**
- `SignatureHelpProviderTests` — verify Roslyn returns signatures for wrapper type methods, verify overload count, verify active parameter index changes with cursor position
- Verify XML doc descriptions appear in Roslyn's signature help results

**AddIn tests (`formula-boss.AddinTests`):**
- Type `table.Rows().Where(` in the editor → verify signature help popup appears
- Type `,` → verify active parameter updates
- Press Up/Down → verify overload cycles
- Press `)` → verify popup dismisses
- Verify popup coexists with completion window (type `table.Rows().Where(r => r.` should show completions while signature help is visible for `Where`)

## Open Questions

None — ready for review.
