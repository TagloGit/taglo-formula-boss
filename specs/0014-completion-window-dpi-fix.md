# 0014 — Completion Window DPI Positioning Fix

## Problem

After a period of usage on a multi-monitor setup, the AvalonEdit `CompletionWindow` (intellisense popup) appears horizontally offset to the left of the expected position. Once the mispositioned popup appears, it captures keystrokes (normal CompletionWindow behavior) so the user cannot type further — only Escape dismisses it, but it reappears broken on the next trigger character. This was previously observed in spec 0009 as "the intellisense popup becomes positionally disconnected from the floating editor during window movement" — 0009 addressed the crash side but not the root cause.

### Root Cause

AvalonEdit's `CompletionWindowBase.UpdatePosition()` calculates screen position using:

```
Point location = textView.PointToScreen(visualLocation);  // physical pixels
// ... boundary checks ...
bounds = bounds.TransformFromDevice(textView);             // convert to DIPs
this.Left = bounds.X;                                      // set WPF position
```

On a Per Monitor Aware V2 thread (as Formula Boss uses), `PointToScreen()` returns physical pixels at the **source monitor's DPI**, but `TransformFromDevice()` converts using the **window's current DPI context**. When these don't match (e.g., editor moved between monitors with different DPI, or after a monitor sleep/wake that changes the effective DPI), the X coordinate is scaled by the wrong factor — producing a horizontal offset.

This is a [known AvalonEdit limitation](https://github.com/icsharpcode/AvalonEdit/issues/198) that remains unfixed.

### Why It Appears to Persist Across Sessions

The bug likely reproduces consistently whenever the monitor arrangement creates a DPI mismatch for the editor window's position. The reinstall may have coincided with a change in monitor state, or the installer re-registers components that reset some DPI caching. The `editor-settings.json` file stores window dimensions but not position, and `_hasBeenPositioned` resets on every process start — so the persistence mechanism is the environment (monitor layout), not saved state.

## Proposed Solution

Replace AvalonEdit's built-in `CompletionWindow` with a custom implementation that uses WPF `Popup` with `PlacementMode.Custom` — the same pattern already proven to work correctly in `SignatureHelpPopup`. The `Popup` approach avoids DPI issues because the custom placement callback uses relative coordinates (subtracting `PointToScreen(origin)` from `PointToScreen(caret)`), which cancels out any DPI scaling.

### Approach

1. Create a `DpiAwareCompletionWindow` that wraps AvalonEdit's `CompletionList` inside a WPF `Popup` instead of inheriting from `CompletionWindowBase` (which is a `Window`).
2. Use `PlacementMode.Custom` with `PlacementTarget = editor.TextArea.TextView`, mirroring `SignatureHelpPopup.PlacePopup()`.
3. Wire up the same keystroke forwarding that `CompletionWindowBase` provides: `TextEntering` for filtering, `TextEntered` for insertion, Escape/Enter/Tab for commit/dismiss.
4. Keep the existing `CompletionList` (which handles filtering, selection, and item rendering) — only replace the window/positioning layer.

### What Changes

- **`FloatingEditorWindow.ShowCompletion()`** — instantiate `DpiAwareCompletionWindow` instead of `CompletionWindow`
- **New `DpiAwareCompletionWindow` class** — Popup-based host for `CompletionList`, handles positioning and keystroke integration
- **`EditorBehaviorHandler`** — may need minor adjustments if the new window's keystroke handling differs

### What Doesn't Change

- `RoslynCompletionProvider`, `CompletionData`, `CompletionHelpers` — completion data pipeline is unchanged
- `SignatureHelpPopup` — already works correctly
- `EditorSettings` — no new settings needed

## User Stories

- As a Formula Boss user on a multi-monitor setup, I want intellisense to appear at the correct position regardless of which monitor the editor is on, so that I can use completions reliably.
- As a Formula Boss user, I want to be able to type continuously with intellisense filtering my input, without the popup blocking my keystrokes due to positioning errors.

## Acceptance Criteria

- [ ] Completion popup appears directly below the caret on single-monitor setups
- [ ] Completion popup appears directly below the caret after moving the editor between monitors with different DPI scales
- [ ] Typing characters filters the completion list (keystrokes not swallowed)
- [ ] Tab/Enter commits the selected completion item
- [ ] Escape dismisses the completion popup
- [ ] Ctrl+Space force-shows completions
- [ ] Completion popup respects screen boundaries (doesn't render off-screen)
- [ ] `CloseWhenCaretAtBeginning` behavior preserved (popup closes if caret moves before the trigger point)
- [ ] Filtering mode works (non-matching items hidden, not just de-prioritized)
- [ ] No regression in SignatureHelpPopup behavior

## Out of Scope

- Upgrading AvalonEdit version (the DPI issue is unfixed upstream)
- Changing the completion data pipeline or Roslyn workspace
- Fixing the `SignatureHelpPopup` (it already works correctly with the Popup pattern)
- General DPI refactoring of other UI components

## Risk: Input Blocking Mechanism Not Fully Understood

The DPI mismatch explains the horizontal offset, but the exact mechanism that blocks keyboard input when the popup is mispositioned is not confirmed. The leading theory is focus theft — `CompletionWindow` is a `Window` that contains a `ListBox`, and if the ListBox captures keyboard focus (instead of the TextArea), regular keystrokes go to the ListBox while Tab still works for selection. Other possibilities include exceptions in AvalonEdit's keystroke handler during `UpdatePosition()`, or a rapid open/close cycle that eats characters in event interleaving.

The Popup-based replacement should fix both issues: `Popup` doesn't participate in WPF's window activation/focus system, so it can't steal focus from the TextArea. **If the Popup replacement fixes the offset but typing is still blocked, investigate the input mechanism explicitly** — add diagnostic logging to `EditorBehaviorHandler.OnTextEntering` and the completion window's keystroke hooks to identify which handler is consuming the keystrokes.

## Open Questions

- **Q1:** Did the reinstall truly fix a persistent bug, or did something else change at the same time (monitor arrangement, docking state, Windows display settings)? Understanding this would confirm whether there's a secondary persistence mechanism we're missing, but it doesn't affect the fix — replacing the Window with a Popup fixes the DPI issue regardless.
- **Q2:** Should we add diagnostic logging for completion window positioning to help debug any future positioning issues? (Lightweight — just log the computed position and current DPI on each show.)
