# 0007 — MVP Launch (v0.1.0)

## Problem

Formula Boss is feature-complete enough for early test users, but the repo isn't ready for public attention. There's no release artifact, no contributor documentation, no branch protection, and the README still has placeholder sections. Before driving traffic to the repo we need to button everything up.

## Proposed Solution

An epic covering four areas: finish pre-launch feature work, harden crash resilience, prepare documentation and repo hygiene, and publish a signed v0.1.0 GitHub Release (64-bit only).

## Pre-Launch Feature Work

These existing backlog issues must be completed before launch:

- **#99 — RowCollection.Aggregate and Scan** — Blocks the "running totals" user journey
- **#100 — Cross-sheet range references** — Removes a common workaround requirement
- **#51 — Crash resilience** — Investigate unhandled exception paths and add top-level guards so the add-in never takes down the Excel process

### Nice-to-Have (not launch blockers)

- #16 — Show structural LET errors with squiggly underlines
- #17 — Auto-format LET formulas with indentation
- #55 — Auto-close function parentheses

## Crash Resilience (#51)

The goal isn't to find every possible crash — it's to ensure that when something unanticipated goes wrong, the add-in fails gracefully rather than taking down Excel. This means:

- Top-level try/catch around all entry points (SheetChange handler, editor events, compilation pipeline)
- AppDomain.UnhandledException / TaskScheduler.UnobservedTaskException handlers that log and swallow rather than propagate
- WPF Dispatcher.UnhandledException handler on the editor thread
- Consider a simple error log file so users can report issues

## Documentation

- **README.md** — Update the License section (currently says `[TBD]`), add a logo/mascot image, add a "Download" section pointing to the GitHub Release
- **CONTRIBUTING.md** — Build instructions, code style conventions (point to .editorconfig), what PRs are welcome, testing requirements
- **Logo export** — Export PNG from `logo-final.html` for use in README

## Distribution

- **Code signing** — Integrate Sectigo certificate into the build or document the manual signing step
- **GitHub Release** — Create a `v0.1.0` release with signed 64-bit `.xll` as an attached asset
- **Release notes** — Brief summary of capabilities (inline expressions, wrapper types, cell formatting access, floating editor)
- **Fresh machine test** — Verify the signed `.xll` installs and works on a clean machine (no dev tools)
- **64-bit only** — Ship 64-bit Excel only for v0.1.0; document the limitation

## Repo Hygiene

- **Secrets audit** — Scan repo history for any leaked secrets, API keys, or personal paths
- **Branch protection on main** — Require PR reviews, prevent force-push
- **Issue templates** — Add bug report and feature request templates
- **Repo metadata** — Add description and topic tags ("excel", "exceldna", "add-in", "csharp", "competitive-excel")
- **Close #23** — Superseded by this epic's child issues

## Branding & First Impression

The README and the add-in itself are the first things users experience — both need to make an impact.

- **README** — Not just informative but impressive. Lead with the logo/mascot, follow with a GIF or short screen recording showing the backtick workflow end-to-end (type expression, see it compile, see the result). The GIF is the single highest-impact asset for the GitHub page. Include a link to the Taglo website.
- **Logo PNG** — Export from `logo-final.html` at appropriate sizes for README and ribbon
- **Add-in ribbon branding** — Add the Formula Boss logo/icon to the Excel ribbon tab
- **Editor branding** — Small status bar at the bottom of the floating editor with Taglo name/logo (subtle, not intrusive)

## Out of Scope

- 32-bit Excel support (requires VM setup; revisit post-launch)
- VBA transpiler (#21), UDF persistence (#20), named UDFs (#19), algorithm library (#22) — post-launch phases
- Launch communications (Reddit, LinkedIn, etc.) — Tim's domain, not tracked in backlog

## Acceptance Criteria

- [ ] #99 and #100 closed
- [ ] #51 closed — crash resilience guards in place
- [ ] README updated with logo, GIF/screen recording, license link, download section
- [ ] CONTRIBUTING.md exists
- [ ] Formula Boss logo/icon on the Excel ribbon tab
- [ ] Floating editor has Taglo-branded status bar
- [ ] Signed 64-bit .xll published as GitHub Release v0.1.0
- [ ] Release verified on a fresh machine
- [ ] Branch protection enabled on main
- [ ] Issue templates added
- [ ] Repo description and topics set
- [ ] No secrets in repo history
- [ ] #23 closed as superseded

## Open Questions

None — all questions resolved during planning discussion.
