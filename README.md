# Formula Boss

An Excel add-in that lets power users write inline C# expressions inside Excel formulas. Expressions operate on typed wrappers around Excel data and compile to UDFs at runtime via ExcelDNA and Roslyn. Built for competitive Excel.

## What It Does

Type a formula with backtick-delimited C# expressions:

```
'=SUM(`data.Cells.Where(c => c.Color == 6).Select(c => c.Value)`)
```

Formula Boss detects the expression, compiles it to a UDF, and rewrites the cell formula automatically.

## Quick Examples

```
# Filter table rows by column values
'=`tblSales.Rows.Where(r => r["Region"] == "EMEA" && r["Revenue"] > 1000)`

# Top 5 by price
'=`tbl.Rows.OrderByDescending(r => r["Price"]).Take(5)`

# Element-wise transform preserving 2D shape
'=`myRange.Map(x => x * 1.1)`

# Multi-input via LET
=LET(maxPop, XLOOKUP(...), result, `tblCountries.Rows.Where(r => r["Population"] < maxPop)`, result)

# Statement blocks for complex logic
'=`{ var total = sales.Sum(); if (total > target) return "Over"; return "Under"; }`
```

## Features

- **Three wrapper types** — `ExcelScalar`, `ExcelArray`, `ExcelTable` with a unified operation API (`.Where()`, `.Select()`, `.Sum()`, `.Rows`, etc.)
- **Automatic parameter detection** — all free variables become UDF parameters; no explicit declaration needed
- **Cell formatting access** — filter/aggregate by color, bold, font via `.Cells`
- **Floating editor** — `Ctrl+Shift+`` ` opens an editor with syntax highlighting, error squiggles, and Roslyn-powered autocomplete
- **LET integration** — backtick expressions work inside `=LET(...)` formulas
- **Range references** — `A1:C10` works directly in expressions alongside named ranges and tables

## Building

Requires .NET 6 SDK.

```bash
dotnet build formula-boss/formula-boss.slnx
dotnet test formula-boss/formula-boss.slnx
```

## Loading in Excel

1. Build the project
2. Open Excel → File → Options → Add-ins → Manage: Excel Add-ins → Go
3. Browse to `formula-boss/bin/Debug/net6.0-windows/formula-boss64.xll` (or `formula-boss.xll` for 32-bit)

## Documentation

- [User Specification](specs/0005-formula-boss-user-spec.md) — full expression language, type system, and operation reference
- [Architecture Specification](specs/0006-formula-boss-architecture.md) — pipeline, runtime, and technical design

## License

[TBD]
