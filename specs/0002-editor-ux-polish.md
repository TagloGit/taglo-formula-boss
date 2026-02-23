# 0002 — Editor UX Polish

## Problem

The Formula Boss floating editor handles the "happy path" of text editing well, but deviates from standard code-editor conventions in many small ways that accumulate into a frustrating experience. Users whose muscle memory comes from Visual Studio / VS Code are repeatedly surprised by missing or incorrect behaviors — particularly around indentation, bracket handling, and cursor positioning after mistakes.

The current implementation adds behaviors incrementally in `EditorBehaviors.cs` on top of AvalonEdit's defaults. This has led to gaps (no smart backspace, no paired delete, no smart Home) and inconsistencies when the user leaves the "perfect" editing flow. Rather than patching individual cases, this spec defines the target UX as a coherent whole so that the implementation can be architected to handle them uniformly.

## Reference Point

Visual Studio C# editing is the benchmark. Where a behavior matches VS, no further justification is needed.

## Proposed Solution

Replace or restructure the editor behavior layer (`EditorBehaviors.cs` and key handling in `FloatingEditorWindow.xaml.cs`) with a unified set of behaviors covering all standard code-editor interactions. AvalonEdit's built-in `CSharpIndentationStrategy` may be retained or replaced depending on what gives the cleanest implementation.

Indent size should be a configurable constant (default: 2 spaces) surfaced in `EditorSettings`.

---

## Behaviors

### B1 — Smart Home Key

**First press:** Cursor moves to the first non-whitespace character on the line.
**Second consecutive press:** Cursor moves to column 0.
**Third press:** Back to first non-whitespace.

With **Shift** held, selection extends accordingly.

### B2 — Smart Backspace in Leading Whitespace

When the cursor is in leading whitespace (only spaces to the left on the current line), Backspace removes characters back to the previous indent stop (nearest multiple of indent size), not one space at a time.

Example (indent size = 2):
- Cursor at column 5 → Backspace moves to column 4
- Cursor at column 4 → Backspace moves to column 2
- Cursor at column 2 → Backspace moves to column 0

### B3 — Auto-Delete Paired Brackets

When the cursor is between an empty auto-inserted pair — `()`, `[]`, `{}`, `""`, ` `` ` — pressing Backspace deletes both the opening and closing character.

### B4 — Smart Auto-Close Suppression for Quotes

Current behavior: typing `"` always inserts a closing `"`. This causes problems when adding a closing quote to existing text.

New behavior: suppress auto-close when the character immediately to the **right** of the cursor is an identifier character (letter, digit, `_`) or another quote. Auto-close only fires when the right side is whitespace, a closing bracket, a comma, a semicolon, or end-of-line.

This matches VS Code's `autoCloseBefore` logic.

### B5 — Surround Selection with Brackets/Quotes

When text is selected and the user types an opening bracket (`(`, `[`, `{`) or quote (`"`, `` ` ``), the selection is wrapped with the pair rather than replaced. Cursor is placed after the closing character; the wrapped text remains selected.

Example: select `foo`, type `(` → `(foo)` with `foo` still selected.

Backtick wrapping produces `` `selection` `` — useful for quickly wrapping text in DSL expression delimiters.

### B6 — Tab / Shift+Tab Indent and Dedent

**Tab, no selection:** Insert spaces to the next tab stop at the cursor position (VS-style).

**Tab, with multiline selection:** Indent all selected lines by one indent unit. Selection is preserved spanning the same lines.

**Shift+Tab (always):** Remove one indent level from the current line (or all selected lines). Works regardless of cursor position on the line. If a line has fewer spaces than one indent unit, remove all leading whitespace.

Selection is preserved after indent/dedent so the user can press Tab/Shift+Tab repeatedly.

### B7 — Enter at Structural Boundaries (Reformat)

When Enter is pressed at specific structural points within a single-line expression, the editor reformats the statement into a correct multi-line layout rather than inserting a raw newline.

**Structural boundaries:**

1. **Between `{` and `}`** (existing, retain): Expand to three lines — opening brace, indented cursor line, closing brace.

2. **After `=>` before `{`** in a lambda: Break the lambda body onto the next line with increased indent.

3. **Before closing `)` or `]`**: Move the closer to a new line outdented to match the opener's indent. Insert a new indented line for the cursor.

4. **Between LET binding pairs**: When Enter is pressed between the value and the next binding name (i.e. after a comma in `name, value, |nextName, ...`), format with each binding on its own line.

**When Enter is pressed at a non-structural point**, fall back to standard behavior: insert a newline and inherit the current line's indentation. The cursor must remain aligned with the content that follows the break point — e.g. pressing Enter before `arg2` in `fn(arg1, |arg2)` should place the cursor directly in front of `arg2` on the new line (with matching indentation), not at column 0.

### B8 — Auto-Indentation on Enter

**After an opening `(`, `[`, `{`:** New line is indented one level deeper than the line containing the opener.

**After a line that doesn't end with an opener:** New line inherits the same indentation as the current line.

**Trailing whitespace cleanup:** If an auto-indented blank line is left empty (user moves away without typing), the trailing whitespace is removed.

### B9 — Cut / Copy Empty Selection = Whole Line

**Ctrl+X with no selection:** Cut the entire current line (including newline).
**Ctrl+C with no selection:** Copy the entire current line.

This matches VS / VS Code behavior.

---

## User Stories

- As a user, I want Home to jump to the start of my code (not column 0) so I can navigate indented formulas quickly.
- As a user, I want Backspace to remove a full indent level so I don't have to press it multiple times in whitespace.
- As a user, I want to delete both brackets when I backspace on an empty `()` so I can quickly undo auto-inserted pairs.
- As a user, I want to add closing quotes without the editor doubling them, so correcting mistakes is painless.
- As a user, I want Tab/Shift+Tab to indent and dedent lines so I can restructure formulas without manual space counting.
- As a user, I want Enter at structural points to reformat my expression correctly so I don't have to manually fix indentation after breaking a line.
- As a user, I want to wrap selected text in brackets by typing `(` so I can add grouping without cut-and-paste.

## Acceptance Criteria

- [ ] B1: Smart Home toggles between first non-whitespace and column 0 (with Shift support)
- [ ] B2: Backspace in leading whitespace removes to previous indent stop
- [ ] B3: Backspace between empty pairs deletes both characters
- [ ] B4: Quote auto-close is suppressed when right-side character is identifier/quote
- [ ] B5: Typing opener with selection wraps the selection
- [ ] B6: Tab indents selected lines; Shift+Tab dedents; selection preserved
- [ ] B7: Enter at structural boundaries reformats to correct multi-line layout
- [ ] B8: Enter after opener indents; otherwise inherits indent; trailing whitespace cleaned
- [ ] B9: Ctrl+X/C with no selection operates on the whole line

## Out of Scope

- Multi-cursor / column selection (future consideration)
- LET formula auto-formatting on load/save (#17 — separate issue)
- LET structural error squiggles (#16 — separate issue)
- Autocomplete / intellisense changes
- Syntax highlighting changes

## Resolved Questions

1. **Enter mid-argument-list** — No special multi-line reformat needed. Standard behavior applies, but cursor must stay aligned in front of the content after the break (see B7 fallback note).
2. **Backtick surround** — Yes, included in B5.
3. **Indent size config** — Stored in `EditorSettings` alongside window size.
