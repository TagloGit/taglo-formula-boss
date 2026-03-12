<p align="center">
  <img src="assets/formula-boss.gif" alt="Formula Boss" width="128" />
</p>

<h1 align="center">Formula Boss</h1>

<p align="center">
  This is the Formula Boss. It wants your formulas. You must feed it. NOW.
</p>
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

**But don't do it like that**:
- Use the visual editor by pressing _Ctrl+Shift+`_.
- Better yet, make your expression a LET argument, so that Formula Boss will give it a nice name.
- Outside of a LET, your function won't be editable - inside a LET, you can use the visual editor to edit your function as much as you like.

🚨 Warning: The Formula Boss is forgetful. It won't remember your functions next time you open Excel. If you used a LET function, you can recreate them, otherwise they're gone for good!

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
- **Floating editor** — `Ctrl+Shift+`` ` opens an editor with syntax highlighting, error squiggles, and autocomplete
- **LET integration** — backtick expressions work inside `=LET(...)` formulas
- **Range references** — `A1:C10` works directly in expressions alongside named ranges and tables

## Demos

### Column name references

Sum of `Price × Volume Sold` from a table — with column-name intellisense that auto-converts to bracket syntax. Shows editing the formula live to add `.Sum()`.

<video src="https://private-user-images.githubusercontent.com/53003551/560283047-9d04fa83-2db5-4cea-af4e-9ca30a71bcb5.mp4" controls width="100%"></video>

### Cell properties

Filter cells by formatting: `steps.Cells.Where(c => c.Bold)`. Toggling bold on cells updates the output in real time. Then editing the formula to handle the empty case with a statement block.

<video src="https://private-user-images.githubusercontent.com/53003551/560281967-4eb56f3b-4d52-4048-bc08-10b7145174ec.mp4" controls width="100%"></video>

### Recamán's sequence

A statement expression with a `foreach` loop, `List<double>`, and visited-step tracking — computing the full Recamán's sequence. Then editing to `.Skip(1)` the starting step.

<video src="https://private-user-images.githubusercontent.com/53003551/560282147-16d3d2cc-38b2-4be2-96fc-1d917a368de7.mp4" controls width="100%"></video>

## Download

**[Download the latest release](https://github.com/TagloGit/taglo-formula-boss/releases/latest)** — runs the installer, no admin rights needed.

Requires **64-bit Excel** (Microsoft 365 or Excel 2019+) on Windows 10/11. The installer bundles the .NET 6 runtime.

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
