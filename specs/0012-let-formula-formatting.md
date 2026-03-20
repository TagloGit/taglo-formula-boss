# 0012 — LET Formula Auto-Formatting

## Problem

LET formulas become hard to read as they grow. The editor currently only formats LET formulas that contain Formula Boss backtick expressions — plain LET formulas are loaded and saved verbatim. Users need consistent, configurable formatting for all LET formulas to maintain readability, especially with nested LETs and many bindings.

## Proposed Solution

Add a general-purpose LET formatter that:
1. Formats **all** LET formulas (not just those with backtick expressions) on load into the editor and on save
2. Respects user-configurable settings for indent size, nesting depth, and line length
3. Preserves content the user has intentionally formatted (non-LET function calls, backtick expression internals)
4. Integrates with the existing `LetFormulaRewriter` so processed formulas also respect formatting settings

## Formatting Rules

### Basic Structure

Each LET binding (name-value pair) appears on its own line, indented:

```
=LET(
    x, A1,
    y, B1,
    x + y)
```

- Closing parenthesis sits on the same line as the result expression (Style A)
- Name and value always appear on the same line: `name, value,`
- Indent uses spaces only (configurable size, default from `EditorSettings.IndentSize`)

### Short-Binding Inlining (Max Line Length)

When `MaxLineLength > 0`, short initial bindings may be inlined on the opening line:

```
=LET(x, A1, y, B1,
    total, SUM(x, y),
    total)
```

Bindings are placed on the `=LET(` line left-to-right until adding the next binding would exceed `MaxLineLength`. Once any binding wraps, all subsequent bindings appear on their own indented lines.

When `MaxLineLength = 0` (default), every binding always wraps — no inlining.

### Nested LET Formatting

Nested LETs are formatted with additional indentation relative to their parent:

```
=LET(
    x, A1,
    result, LET(
        y, B1,
        z, C1,
        y + z),
    x + result)
```

The `NestedLetDepth` setting controls how many levels deep to format:
- `1` (default): Only the top-level LET is formatted; nested LETs are left as-is
- `2`: Top-level and one level of nesting are formatted
- `0`: No LET formatting at all (auto-formatting disabled for LETs)

### Non-LET Content Preservation

- **Non-LET function calls** within binding values are never reformatted. `FILTER(SORT(A1:A100, 2, -1), B1:B100 > 0)` stays on one line.
- **Backtick expression content** is never reformatted — left exactly as the user wrote it.
- **String literals** (`"hello, world"`) are parsed correctly and not split on internal commas.
- **Array constants** (`{1,2,3;4,5,6}`) are kept on one line.

### LET Inside Other Functions

When a LET appears inside another function (e.g. `=SUMPRODUCT(LET(x, A1:A10, x * 2))`), the LET is formatted according to the rules and the wrapping function is preserved around it:

```
=SUMPRODUCT(LET(
    x, A1:A10,
    x * 2))
```

If this proves too complex to implement reliably, the fallback is: **no formatting for formulas that don't start with `=LET(`**. This can be revisited later.

## User Stories

- As a power user with large LET formulas, I want them automatically formatted when I open the editor so I can quickly read and understand the structure.
- As a user with specific formatting preferences, I want to configure indent size and nesting depth so the formatting matches my style.
- As a user who writes compact formulas, I want to inline short bindings on one line so simple LETs don't take up too much vertical space.

## Settings

New settings in `EditorSettings`:

| Setting | Type | Default | Description |
|---|---|---|---|
| `AutoFormatLet` | `bool` | `true` | Enable/disable LET auto-formatting |
| `NestedLetDepth` | `int` | `1` | How many levels of nested LETs to format (0 = off) |
| `MaxLineLength` | `int` | `0` | Max line length before wrapping (0 = always wrap) |

`IndentSize` already exists in `EditorSettings` (default 2).

Note: `AutoFormatLet = false` disables the formatter entirely. `NestedLetDepth = 0` with `AutoFormatLet = true` would be contradictory — if `AutoFormatLet` is false, `NestedLetDepth` is ignored.

## When Formatting Applies

1. **On load into editor** — formula is formatted before display
2. **On save from editor** — formula is formatted before writing to cell

No on-the-fly formatting while the user types. A manual "format" action is unnecessary since open-then-save achieves the same result.

## Acceptance Criteria

- [ ] All LET formulas are formatted on editor load, not just backtick formulas
- [ ] All LET formulas are formatted on editor save
- [ ] Formatting respects `IndentSize` setting (existing)
- [ ] `AutoFormatLet` setting toggles formatting on/off
- [ ] `NestedLetDepth` controls depth of nested LET formatting
- [ ] `MaxLineLength` enables short-binding inlining when > 0
- [ ] Settings are exposed in the Settings dialog
- [ ] String literals with commas are parsed correctly (not split)
- [ ] Array constants are preserved on one line
- [ ] Non-LET function calls within bindings are not reformatted
- [ ] Backtick expression content is preserved exactly
- [ ] `LetFormulaRewriter` uses formatting settings instead of hardcoded indent
- [ ] Existing tests updated/passing; new tests for formatting edge cases

## Out of Scope

- Formatting non-LET functions (e.g. wrapping long `IF` chains)
- On-the-fly formatting as the user types
- Tabs as indent style
- Configurable closing parenthesis placement
- Separate-line name/value pairs (`name,\n    value,`)

## Open Questions

- Should the settings UI group formatting options under a "Formatting" section in the Settings dialog, or keep them flat alongside existing settings?
