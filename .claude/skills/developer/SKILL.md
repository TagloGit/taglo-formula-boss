---
name: developer
description: "Execute approved implementation plans for Formula Boss. Usage: /developer <issue-number>"
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# Developer — Formula Boss

Execute approved implementation plans for the Formula Boss Excel add-in.

## Instructions

### Starting Work
1. Read the issue: `gh issue view <number> -R TagloGit/taglo-formula-boss`
2. Check for linked spec/plan docs in the issue body
3. Read spec and plan files if referenced
4. Update issue status:
   ```bash
   gh issue edit <number> -R TagloGit/taglo-formula-boss --remove-label "status: backlog" --add-label "status: in-progress"
   ```
5. Create a branch: `git checkout -b issue-<number>-short-description`

### Working
- Make small, well-described commits
- Follow repo coding standards (see below)
- Build: `dotnet build formula-boss/formula-boss.slnx`
- Test: `dotnet test formula-boss/formula-boss.slnx`

### When Blocked
- Add `blocked: tim` label: `gh issue edit <number> -R TagloGit/taglo-formula-boss --add-label "blocked: tim"`
- Add a comment explaining what you need
- Stop work on this issue — do not stall silently

### When You Discover Unplanned Work
- Stop — don't absorb it into the current task
- WIP commit: `WIP: #<number> - paused for new issue`
- Create a new issue for the blocker with `status: backlog`
- Comment on the original issue: "Blocked by #N"
- Ask Tim whether to switch or park it

### Raising a PR
1. Push branch: `git push -u origin <branch-name>`
2. Create PR with `Closes #<number>` in body
3. Update issue: remove `status: in-progress`, add `status: in-review`

## Repo-Specific Guidance

- .NET 6 / C# 10, ExcelDNA 1.9
- File-scoped namespaces, `_camelCase` private fields, nullable reference types enabled
- **Critical:** Never reference ExcelDNA types directly in generated code — use reflection (see CLAUDE.md for details)
- After significant refactors, remind Tim to check ReSharper warnings
- Specs: `specs/`, Plans: `plans/`

## Self-Improvement

When you notice a recurring problem, a workflow gap, or something that would help future instances:
1. Create a `process` issue on TagloGit/taglo-pm describing the observation and suggested improvement
2. Reference the specific skill/file that should change (if known)
3. Continue your current work — don't block on the improvement
