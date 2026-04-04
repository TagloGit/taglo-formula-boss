# 0014 — Completion Window DPI Fix — Implementation Plan

## Overview

Replace AvalonEdit's `CompletionWindow` (a `Window` with DPI-broken positioning) with a custom `CompletionPopup` that hosts AvalonEdit's `CompletionList` inside a WPF `Popup`. The `Popup` uses `PlacementMode.Custom` with relative coordinates — the same pattern proven in `SignatureHelpPopup` — which eliminates the DPI mismatch.

AvalonEdit's `CompletionList` is reusable standalone. It handles filtering, selection, item rendering, and keyboard navigation (Up/Down/PageUp/PageDown/Home/End/Tab/Enter). We only need to replace the positioning and event-wiring layer that `CompletionWindowBase` provides.

## Files to Touch

- **`formula-boss/UI/Completion/CompletionPopup.cs`** (new) — Popup-based completion host, replaces `CompletionWindow`
- **`formula-boss/UI/FloatingEditorWindow.xaml.cs`** — Switch from `CompletionWindow` to `CompletionPopup`
- **`formula-boss/UI/EditorBehaviorHandler.cs`** — Route navigation keys to `CompletionPopup` when open

## Design: CompletionPopup

### Hosting

- WPF `Popup` with `PlacementMode.Custom`, `PlacementTarget = TextArea.TextView`
- `CustomPopupPlacementCallback` uses the same relative-coordinate technique as `SignatureHelpPopup.PlacePopup()`: subtract `PointToScreen(origin)` from `PointToScreen(caret)`, producing DPI-neutral offsets
- Place below the caret line; if insufficient space, place above
- `StaysOpen = true` (we manage closing ourselves)
- `AllowsTransparency = true` (needed for border styling)

### Content

- Hosts AvalonEdit's `CompletionList` control directly
- Wrap in a `Border` for consistent styling (match current completion window appearance)
- Set `CompletionList.IsFiltering = true`
- Style the `ListBox` the same way `FloatingEditorWindow.ShowCompletion()` currently does (highlight brush, disabled horizontal scrollbar)

### Offset Tracking (from CompletionWindowBase)

- `StartOffset` / `EndOffset` — int properties tracking the completion region in the document
- Subscribe to `TextArea.Document.Changing` to update offsets as the document changes:
  - Insertions at or after StartOffset: extend EndOffset
  - Removals before StartOffset: shift both offsets
  - Removal that deletes into the region before StartOffset: close

### Filtering (from CompletionWindow)

- Subscribe to `TextArea.Caret.PositionChanged`
- Extract text between `StartOffset` and `CaretOffset`
- Call `CompletionList.SelectItem(text)` to filter
- If `CaretOffset < StartOffset` (or `== StartOffset` when `CloseWhenCaretAtBeginning`): close
- If `CaretOffset > EndOffset`: close

### Keystroke Handling

- Subscribe to `TextArea.PreviewKeyDown` (or use a `TextAreaStackedInputHandler`)
- Route these keys to `CompletionList.HandleKey()`:
  - **Up, Down, PageUp, PageDown, Home, End** — list navigation
  - **Tab, Enter** — commit (triggers `InsertionRequested`)
  - **Escape** — close popup
- `CompletionList.InsertionRequested` handler: close popup first, then call `item.Complete()` on the selected item (close-before-complete order is critical per AvalonEdit's design)

### Focus / Lifecycle

- Subscribe to `TextArea.LostKeyboardFocus` → close (async dispatch, check if TextArea is still focused)
- Subscribe to `TextArea.DocumentChanged` → close immediately
- Subscribe to parent window `LocationChanged` → force Popup reposition (`HorizontalOffset = 0` trick, same as SignatureHelpPopup)
- Provide a `Closed` event so `FloatingEditorWindow` can null out its reference and reset `IsBracketContext`

### Tooltip

- AvalonEdit's `CompletionWindow` shows a tooltip for the selected item's `Description`. Our `CompletionData` currently returns a nullable string description.
- Add a `ToolTip` positioned to the right of the popup, updated on `CompletionList.SelectionChanged`

### Public API

```csharp
internal sealed class CompletionPopup
{
    public CompletionPopup(TextArea textArea);
    public CompletionList CompletionList { get; }
    public int StartOffset { get; set; }
    public bool CloseWhenCaretAtBeginning { get; set; }
    public bool IsOpen { get; }
    public event EventHandler? Closed;
    public void Show();
    public void Close();
}
```

## Order of Operations

### Step 1: Create CompletionPopup

New file `formula-boss/UI/Completion/CompletionPopup.cs`:
- Popup hosting CompletionList
- PlacementMode.Custom with DPI-safe positioning callback
- Document.Changing subscription for offset tracking
- Caret.PositionChanged subscription for filtering
- PreviewKeyDown routing to CompletionList.HandleKey()
- InsertionRequested → close + complete
- Focus loss detection
- Tooltip on selection change

This is the bulk of the work. It replicates CompletionWindowBase's event wiring without the Window/positioning code.

### Step 2: Update FloatingEditorWindow

In `ShowCompletion()`:
- Replace `new CompletionWindow(FormulaEditor.TextArea)` with `new CompletionPopup(FormulaEditor.TextArea)`
- Remove the styling code that reaches into `_completionWindow.CompletionList.ListBox` — move that into CompletionPopup's constructor
- Keep the `MinWidth`/`MaxWidth`/`SizeToContent` setup (adapt to Popup's content Border)
- `_completionWindow.Show()` → `_completionPopup.Show()`
- `_completionWindow.Closed` → `_completionPopup.Closed`
- `_completionWindow.CompletionList.IsFiltering = true` → set in CompletionPopup constructor
- `_completionWindow.CompletionList.SelectItem(prefix)` → `_completionPopup.CompletionList.SelectItem(prefix)`

Update all references: rename `_completionWindow` to `_completionPopup` (type changes from `CompletionWindow?` to `CompletionPopup?`).

### Step 3: Update EditorBehaviorHandler

The `IsCompletionWindowOpen` callback is already used to gate Tab handling and Up/Down overload cycling. This continues to work — `FloatingEditorWindow` just returns `_completionPopup != null` instead of `_completionWindow != null`.

**However**, CompletionWindowBase's stacked input handler currently intercepts Up/Down/PageUp/PageDown before they reach `EditorBehaviorHandler.OnPreviewKeyDown`. With that gone, we need CompletionPopup to handle those keys itself via its own PreviewKeyDown subscription on the TextArea. The EditorBehaviorHandler doesn't need to change for this — CompletionPopup subscribes independently.

The one thing to verify: the event ordering. CompletionPopup's PreviewKeyDown handler must run before EditorBehaviorHandler's for Up/Down keys (so the list navigates instead of the caret moving). Since Preview events use tunneling (parent → child), the subscription order on the TextArea matters. CompletionPopup should subscribe when opened and unsubscribe when closed, and its handler should set `e.Handled = true` for navigation keys.

### Step 4: Test

- Build and run AddIn tests
- Manual testing on multi-monitor setup (different DPI scales)
- Test all acceptance criteria from the spec

## Testing Approach

- **AddIn tests**: existing completion-related tests should pass unchanged (they test the data pipeline, not the popup)
- **Manual testing**: the core value is multi-monitor DPI behavior, which can't be automated
  - Open editor on monitor A (e.g. 100% DPI), trigger completion, verify position
  - Move editor to monitor B (e.g. 150% DPI), trigger completion, verify position
  - Type characters with completion open — verify filtering works
  - Tab/Enter to commit, Escape to dismiss
  - Ctrl+Space to force-show
  - Move parent window while completion is open — verify popup follows

## Open Questions

- **Q1:** Should CompletionPopup use a `TextAreaStackedInputHandler` (like CompletionWindowBase) or a direct `PreviewKeyDown` subscription on the TextArea? The stacked handler is AvalonEdit's intended mechanism and handles event ordering cleanly, but adds complexity. A direct PreviewKeyDown subscription is simpler and matches how EditorBehaviorHandler works. **Recommendation:** Start with direct PreviewKeyDown; switch to stacked handler only if we hit event-ordering issues.
