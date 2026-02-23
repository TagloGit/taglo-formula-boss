# 0002 — Editor UX Polish — Implementation Plan

## Overview

Rewrite the editor behavior layer to implement all 9 behaviors from spec 0002. The current `EditorBehaviors.cs` (static helpers) and key handling in `FloatingEditorWindow.xaml.cs` will be restructured into a single cohesive class that owns all editing interactions. AvalonEdit's `CSharpIndentationStrategy` will be replaced with a custom strategy.

No test project exists currently. A test project will be created to cover the new behavior class, since these are pure document-manipulation methods that can be tested without a running WPF application.

## Architecture Approach

**New class: `EditorBehaviorHandler`** — an instance class that takes a `TextEditor` reference and wires up all event handlers internally. This replaces both the static `EditorBehaviors` class and the event handler code in `FloatingEditorWindow.xaml.cs`.

Benefits:
- Single place for all editing behavior (easier to reason about interaction between behaviors)
- Can hold state (e.g. "last Home press was smart-home" for toggle detection)
- Testable via AvalonEdit's `TextDocument` (the document model works without WPF visuals)

**`EditorSettings` gains `IndentSize` property** (default: 2).

## Files to Touch

| File | Change |
|---|---|
| `UI/EditorBehaviors.cs` | **Delete** — replaced by `EditorBehaviorHandler` |
| `UI/EditorBehaviorHandler.cs` | **New** — all 9 behaviors |
| `UI/FloatingEditorWindow.xaml.cs` | **Simplify** — remove all `OnTextEntering`/`OnTextEntered`/`OnPreviewKeyDown` behavior logic; delegate to `EditorBehaviorHandler`. Keep completion window management. |
| `UI/EditorSettings.cs` | **Add** `IndentSize` property (default: 2) |
| `formula-boss.Tests/formula-boss.Tests.csproj` | **New** — test project |
| `formula-boss.Tests/EditorBehaviorHandlerTests.cs` | **New** — unit tests |
| `formula-boss/formula-boss.slnx` | **Update** — add test project |

## Order of Operations

### PR 1 — Infrastructure: Test project + `EditorSettings.IndentSize`

1. Create `formula-boss.Tests` xUnit project referencing the main project and AvalonEdit
2. Add `IndentSize` property to `EditorSettings` (default: 2)
3. Update `FloatingEditorWindow` constructor to use `_settings.IndentSize` instead of hardcoded `4`
4. Verify build and that existing behavior is unchanged

### PR 2 — Core refactor: `EditorBehaviorHandler` with existing behaviors

1. Create `EditorBehaviorHandler` class that takes `TextEditor` + `EditorSettings`
2. Move all existing behavior from `EditorBehaviors` (static) into `EditorBehaviorHandler` (instance methods)
3. Move event handler logic from `FloatingEditorWindow.xaml.cs` into `EditorBehaviorHandler`
   - `EditorBehaviorHandler` subscribes to `TextEntering`, `TextEntered`, `PreviewKeyDown` internally
   - Expose events/callbacks for things the window still needs (completion triggers, formula-apply)
4. Delete `EditorBehaviors.cs`
5. Update `FloatingEditorWindow` to create `EditorBehaviorHandler` and wire up callbacks
6. Write tests for all migrated behaviors (brace expansion, closing-char skip, auto-insert, de-indent)
7. Verify existing behavior is preserved

### PR 3 — New behaviors: B1 (Smart Home), B2 (Smart Backspace), B3 (Paired Delete)

These three are independent, purely local behaviors with no interaction between them.

1. **B1 — Smart Home:** Intercept `Key.Home` in `PreviewKeyDown`. Track whether last action was a Home press to toggle between first-non-whitespace and column 0. Handle `Shift+Home` for selection extension.
2. **B2 — Smart Backspace:** Intercept `Key.Back` in `PreviewKeyDown` when cursor is in leading whitespace. Calculate previous indent stop, remove spaces back to it.
3. **B3 — Paired Delete:** Intercept `Key.Back` when cursor is between an empty pair (`()`, `[]`, `{}`, `""`, ` `` `). Delete both characters.
4. Write tests for each.

### PR 4 — New behaviors: B4 (Smart Quote Suppression), B5 (Surround Selection)

1. **B4:** Modify auto-insert logic — before inserting closing quote, check the character to the right. Suppress if it's an identifier char, digit, `_`, or quote.
2. **B5:** In `TextEntering`, when there's a selection and the typed char is an opener (`(`, `[`, `{`, `"`, `` ` ``), wrap the selection with the pair instead of replacing it. Preserve selection on the wrapped content.
3. Write tests.

### PR 5 — New behaviors: B6 (Tab/Shift+Tab Indent/Dedent)

1. Intercept `Key.Tab` in `PreviewKeyDown`.
   - No selection: insert spaces to next tab stop (VS-style).
   - Multiline selection: indent all selected lines by one indent unit; preserve selection.
2. Intercept `Shift+Tab`: dedent current line or all selected lines by one indent unit. Preserve selection.
3. Write tests.

### PR 6 — New behaviors: B7 (Structural Enter Reformat) + B8 (Auto-Indent on Enter)

This is the most complex PR. It replaces `CSharpIndentationStrategy` with custom Enter handling.

1. **B8 — Auto-Indent on Enter (base case):** On Enter, insert newline + indentation. If previous line ends with `(`, `[`, `{`: indent one level deeper. Otherwise: inherit current line's indent. Cursor placed at end of indentation on new line.
2. **B7 — Structural Enter at boundaries:** Extend Enter handling for:
   - Between `{|}`: existing brace-block expansion (migrated from PR 2)
   - Before `)` or `]`: existing closer expansion (migrated from PR 2)
   - After `=>` before `{`: break lambda body to next line with indent
   - Between LET bindings (after `,` between value and next name): each binding on own line
3. **B7 fallback:** When Enter is at a non-structural point, cursor must remain aligned in front of the content after the break. This means: take the text after the cursor on the current line, move it to the new line with appropriate indentation, and place the cursor at the start of that moved text.
4. **Trailing whitespace cleanup:** Register a handler that trims trailing whitespace from a line when the caret leaves it (if the line is blank).
5. Remove `CSharpIndentationStrategy` — all indentation is now handled by `EditorBehaviorHandler`.
6. Write tests for each structural case and the fallback.

### PR 7 — B9 (Cut/Copy Empty Selection = Whole Line)

1. Intercept `Ctrl+X` / `Ctrl+C` in `PreviewKeyDown` when selection is empty.
2. Select the full line (including newline), then perform the cut/copy via `Clipboard`.
3. Write tests.

## Testing Approach

**Test harness:** Create an AvalonEdit `TextEditor` instance (or just `TextDocument` where possible) in tests. `EditorBehaviorHandler` methods that manipulate the document and caret can be called directly. For key-event-driven behaviors, call the handler methods directly rather than simulating WPF key events.

**Test structure:** One test class per PR / behavior group. Tests assert:
- Document text after the operation
- Caret position after the operation
- Selection state where relevant (B5, B6)

**Test project:** xUnit + FluentAssertions. Reference `ICSharpCode.AvalonEdit` NuGet for the `TextDocument` / `TextEditor` types.

**Note:** AvalonEdit's `TextEditor` requires a WPF dispatcher. Tests may need `[STAThread]` or a test helper that creates a dispatcher. If this proves impractical, extract pure document-manipulation logic into methods that take `TextDocument` + offset and return the new text + offset, tested without WPF dependencies.

## Open Questions

1. **Completion callback design:** `EditorBehaviorHandler` needs to notify `FloatingEditorWindow` to show completions (on `.`, `[`, letter typed) and to apply the formula (on Ctrl+Enter). Options: (a) events, (b) `Action` callbacks passed in constructor, (c) interface. Suggest (b) for simplicity — e.g. `Action onRequestCompletion`, `Action<string> onFormulaApply`.
2. **WPF test feasibility:** Need to verify that AvalonEdit `TextEditor` can be instantiated in a test runner. If not, the pure-logic extraction fallback is straightforward but slightly more work.
