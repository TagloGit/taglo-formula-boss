# Contributing to Formula Boss

Thanks for your interest in contributing! This guide covers what you need to get started.

## Prerequisites

- [.NET 6 SDK](https://dotnet.microsoft.com/download/dotnet/6.0)
- Excel (for integration tests)
- Windows (ExcelDNA is Windows-only)

## Building

```bash
dotnet build formula-boss/formula-boss.slnx
```

## Running Tests

```bash
# Unit and integration tests
dotnet test formula-boss/formula-boss.slnx
```

The `formula-boss.AddinTests` project runs tests with the add-in inside a live Excel instance. These tests are the ultimate proof of correct operation and require Excel to be installed.

## Code Style

Style is enforced via [`.editorconfig`](.editorconfig) and [`Directory.Build.props`](formula-boss/Directory.Build.props):

- File-scoped namespaces
- Private fields: `_camelCase`
- Nullable reference types enabled
- Modern C# patterns preferred (pattern matching, null coalescing, target-typed new)
- Allman brace style (braces on new lines)

Configure your editor to respect the `.editorconfig` and you'll stay consistent.

## Branch and Commit Conventions

- Branch from `main`
- Branch naming: `issue-<number>-short-description`
- Keep commits small and well-described
- Reference issues in PR descriptions with `Closes #<number>`

## What PRs Are Welcome

- Bug fixes with a clear reproduction case
- Improvements to the expression language or wrapper types
- Editor UX enhancements
- Documentation improvements
- Test coverage additions

For larger features, please open an issue first to discuss the approach.

## Testing Requirements

- Unit tests for new pipeline/compiler logic
- AddIn tests for any change to UDF behaviour or Excel integration
- All existing tests must pass before submitting a PR
