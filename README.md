<p align="center">
  <img src="assets/logo-256.png" alt="Formula Boss logo" width="128" />
</p>

<h1 align="center">Formula Boss</h1>

<p align="center">
  Write C# expressions directly in Excel formulas. Built for competitive Excel.
</p>

<p align="center">
  <a href="#demos">Demos</a> &middot;
  <a href="#quick-examples">Examples</a> &middot;
  <a href="specs/0005-formula-boss-user-spec.md">Full Spec</a> &middot;
  <a href="https://www.taglo.io">Taglo</a>
</p>

---

## How It Works

Type a formula with backtick-delimited C# expressions:

```
'=SUM(`data.Cells.Where(c => c.Color == 6).Select(c => c.Value)`)
```

Formula Boss detects the expression, compiles it to a UDF via Roslyn, and rewrites the cell formula automatically.

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

## Demos

### Column name references

Sum of `Price × Volume Sold` from a table — with column-name intellisense that auto-converts to bracket syntax. Shows editing the formula live to add `.Sum()`.

<video src="assets/demo-column-references.mp4" controls width="100%"></video>

### Cell properties

Filter cells by formatting: `steps.Cells.Where(c => c.Bold)`. Toggling bold on cells updates the output in real time. Then editing the formula to handle the empty case with a statement block.

<video src="assets/demo-cell-properties.mp4" controls width="100%"></video>

### Recamán's sequence

A statement expression with a `foreach` loop, `List<double>`, and visited-step tracking — computing the full Recamán's sequence. Then editing to `.Skip(1)` the starting step.

<video src="assets/demo-recaman-sequence.mp4" controls width="100%"></video>

## Download

<!-- TODO: Add link to GitHub Release once published -->

No release is available yet. To try Formula Boss, build from source (see below).

## Building

Requires [.NET 6 SDK](https://dotnet.microsoft.com/download/dotnet/6.0).

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

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for build instructions, code style, and PR guidelines.

## License

[MIT](LICENSE)

---

<sub>Formula Boss is a [Taglo](https://www.taglo.io) project.</sub>
