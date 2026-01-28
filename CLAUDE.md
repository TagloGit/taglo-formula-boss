# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Formula Boss is an Excel add-in that allows power users to write inline expressions using a concise DSL that transpiles to C# UDFs at runtime via ExcelDNA and Roslyn. The primary use case is competitive Excel environments.

## Build Commands

```bash
# Build the solution (requires .NET 10 SDK)
dotnet build formula-boss/formula-boss.slnx

# Run tests (when tests exist)
dotnet test formula-boss/formula-boss.slnx
```

## Architecture

**Tech Stack (proposed):**
- .NET 10 / C# 14
- ExcelDNA - Native C API integration for high-performance UDFs
- Roslyn (Microsoft.CodeAnalysis.CSharp) - Runtime C# compilation
- Microsoft.Office.Interop.Excel - Excel object model access for cell properties (color, formatting)
- AvalonEdit - Floating editor UI (post-MVP)

## Code Style

Enforced via `.editorconfig` and `Directory.Build.props`:
- File-scoped namespaces
- Private fields: `_camelCase`
- Nullable reference types enabled
- Modern C# patterns preferred (pattern matching, null coalescing, target-typed new)
- Root namespace: `FormulaBoss`

## Warnings

The user monitors warnings using ReSharper in Visual Studio. After significant code refactors, ask the user to review ReSharper warnings before committing.

## Specifications

Detailed specifications and implementation plan are in `docs/specs and plan/`:
- `excel-udf-addin-spec.md` - Full DSL specification, user journeys, error handling
- `excel-udf-addin-implementation-plan.md` - Technical stack rationale, 14 implementation phases, risk register
