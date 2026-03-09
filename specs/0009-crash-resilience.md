# 0009 — Crash Resilience

## Problem

Formula Boss runs in-process with Excel. An unhandled exception on any thread — particularly the WPF editor thread — terminates the entire Excel process. This is unacceptable for competitive Excel use where losing a workbook mid-session is catastrophic.

The only observed crash has been a WPF issue: the intellisense popup becomes positionally disconnected from the floating editor during window movement, followed shortly by a fatal crash. No other crashes or hangs have been observed.

The goal is targeted hardening of the known risk areas, not blanket global exception handling. We want to prevent crashes without introducing hangs or leaving the add-in in an unknown state.

## Proposed Solution

Three components:

1. **WPF Dispatcher.UnhandledException handler** — catch and swallow exceptions on the editor thread so WPF bugs don't take down Excel.
2. **TaskScheduler.UnobservedTaskException handler** — prevent fire-and-forget async tasks (e.g., workspace warm-up) from crashing the process.
3. **Simple log file** — write caught exceptions and add-in lifecycle events to a log file for diagnostics.

### What's NOT in scope

- `AppDomain.UnhandledException` — in .NET 6 this is a notification, not a recovery mechanism. The process still terminates. Not useful for crash prevention.
- Timeouts on Roslyn operations — no hangs have been observed. Can be added later if needed.
- Circuit breaker / degraded mode — over-engineering for the current failure profile.
- Out-of-process compilation — massive architectural change not justified by a single WPF positioning bug.

## Detailed Design

### WPF Dispatcher.UnhandledException

When the WPF editor thread's dispatcher is created (in `ShowFloatingEditorCommand`), attach an `UnhandledException` handler:

```csharp
dispatcher.UnhandledException += (sender, e) =>
{
    Logger.Error("WPF dispatcher", e.Exception);
    e.Handled = true; // Prevent process termination
};
```

Setting `e.Handled = true` swallows the exception and keeps the dispatcher running. The editor may be in a visually inconsistent state, but Excel survives. The user can close and reopen the editor to recover.

### TaskScheduler.UnobservedTaskException

In `AddIn.AutoOpen`, register a handler:

```csharp
TaskScheduler.UnobservedTaskException += (sender, e) =>
{
    Logger.Error("Unobserved task", e.Exception);
    e.SetObserved(); // Prevent process termination
};
```

This catches exceptions from fire-and-forget tasks that would otherwise crash the finalizer thread.

### Log File

- **Location:** `%LOCALAPPDATA%\FormulaBoss\logs\formulaboss.log`
- **What to log:**
  - Add-in lifecycle: AutoOpen, shutdown
  - Caught exceptions from the guard handlers above
  - Existing entry-point catch blocks (FormulaInterceptor, DynamicCompiler) — add logging to these
- **Format:** Simple timestamped text lines: `[2026-03-09 14:30:05] [ERROR] WPF dispatcher: NullReferenceException at ...`
- **Rotation:** Single file, truncated on add-in load if over 1 MB. No complex rotation scheme.
- **Implementation:** A static `Logger` class with `Info(string message)` and `Error(string source, Exception ex)` methods. Writes are synchronised with a lock. File I/O failures are silently ignored (the logger must never itself cause a crash).

## Acceptance Criteria

- [ ] WPF editor thread has a `Dispatcher.UnhandledException` handler that logs and sets `Handled = true`
- [ ] `TaskScheduler.UnobservedTaskException` handler registered in AutoOpen that logs and calls `SetObserved()`
- [ ] Static `Logger` class writes to `%LOCALAPPDATA%\FormulaBoss\logs\formulaboss.log`
- [ ] AutoOpen and shutdown are logged
- [ ] Existing catch blocks in FormulaInterceptor and DynamicCompiler log to the file
- [ ] Log file truncated on startup if over 1 MB
- [ ] Verified: injecting a throw in a WPF event handler doesn't crash Excel (manual test)

## Out of Scope

- Roslyn operation timeouts
- Global AppDomain crash handling
- Circuit breaker / auto-disable
- Structured logging frameworks (Serilog, NLog, etc.) — a simple static class is sufficient
