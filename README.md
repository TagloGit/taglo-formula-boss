# Formula Boss

An Excel add-in that allows power users to write inline expressions using a concise DSL that transpiles to C# UDFs at runtime via ExcelDNA and Roslyn.

## Status: MVP Complete

The core pipeline is functional:
- Backtick expressions in formulas are detected and processed
- DSL is parsed, transpiled to C#, compiled at runtime, and registered as UDFs
- Formulas are automatically rewritten to call the generated UDFs

## Quick Start

### Building

Requires .NET 6 SDK.

```bash
dotnet build formula-boss/formula-boss.slnx
dotnet test formula-boss/formula-boss.slnx
```

### Loading in Excel

1. Build the project
2. Open Excel
3. File > Options > Add-ins > Manage: Excel Add-ins > Go
4. Browse to `formula-boss/bin/Debug/net6.0-windows/formula-boss64.xll` (or `formula-boss.xll` for 32-bit Excel)

### Usage

1. Create a named range in Excel (e.g., select `A1:A100`, then type `data` in the Name Box)
2. Type a formula with backtick expressions. The formula must start with `'=` (quote prefix):

```
'=SUM(`data.where(v => v > 0)`)
```

When you press Enter, Formula Boss:
1. Detects the backtick expression
2. Parses and compiles it to a UDF
3. Rewrites your formula to: `=SUM(__udf_abc123(data))`

**Note:** Currently only named ranges work. Cell references like `A1:A10` are not yet supported.

## DSL Syntax

### Data Access

| Syntax | Description |
|--------|-------------|
| `range.values` | Iterate over cell values only (fast path) — **implicit if omitted** |
| `range.cells` | Iterate over cells with object model access (slower, but gives access to formatting) |

**Note:** `.values` is implicit. `data.where(...)` is equivalent to `data.values.where(...)`. Use `.cells` only when you need access to cell formatting properties like `color`, `bold`, etc.

### Methods

| Method | Description |
|--------|-------------|
| `.where(predicate)` | Filter items matching predicate |
| `.select(transform)` | Transform each item |
| `.toArray()` | Materialize results as array — **implicit for collection results** |
| `.sum()` | Sum numeric values |
| `.avg()` / `.average()` | Average of numeric values |
| `.min()` / `.max()` | Minimum/maximum value |
| `.count()` | Count of items |
| `.first()` / `.last()` | First/last item |
| `.orderBy(selector)` | Sort ascending |
| `.orderByDesc(selector)` | Sort descending |
| `.take(n)` / `.skip(n)` | Take/skip n items |
| `.distinct()` | Remove duplicates |

### Cell Properties (requires `.cells`)

| Property | Description |
|----------|-------------|
| `c.value` | Cell value |
| `c.color` | Interior color index |
| `c.rgb` | Interior color as RGB integer |
| `c.row` | Row number |
| `c.col` | Column number |
| `c.bold` | Font is bold |
| `c.italic` | Font is italic |
| `c.fontSize` | Font size |
| `c.format` | Number format string |
| `c.formula` | Cell formula |
| `c.address` | Cell address |

### Operators

- Comparison: `==`, `!=`, `>`, `<`, `>=`, `<=`
- Logical: `&&`, `||`, `!`
- Arithmetic: `+`, `-`, `*`, `/`

### Examples

First, create a named range called `data` pointing to your data range (e.g., `A1:A100`).

```
# Sum positive values (implicit .values)
'=SUM(`data.where(v => v > 0)`)

# Filter by cell color (yellow = index 6)
'=`data.cells.where(c => c.color == 6).select(c => c.value)`

# Get values from bold cells
'=`data.cells.where(c => c.bold).select(c => c.value)`

# Top 5 values
'=`data.orderByDesc(v => v).take(5)`

# Count cells matching condition
'=`data.where(v => v > 50 && v < 100).count()`

# Explicit syntax still works
'=`data.values.where(v => v > 0).toArray()`
```

## Architecture

```
User types formula with backticks
         |
         v
  FormulaInterceptor (SheetChange event)
         |
         v
  BacktickExtractor (extract expressions)
         |
         v
  FormulaPipeline (orchestration)
         |
    +----+----+
    |         |
    v         v
  Lexer    Parser --> AST
              |
              v
      CSharpTranspiler --> C# source
              |
              v
      DynamicCompiler (Roslyn) --> Assembly
              |
              v
      ExcelDNA Registration --> UDF available
              |
              v
      Formula rewritten to call UDF
```

## Known Limitations

1. **Named ranges only**: Cell references like `A1:A10` don't work yet - the parser doesn't handle `:`. Create a named range first.
2. **Quote prefix required**: Formulas must start with `'=` to prevent Excel from evaluating before interception
3. **No IntelliSense**: The DSL editor is just Excel's formula bar - no autocomplete or syntax highlighting yet
4. **Object model is slow**: Using `.cells` accesses the Excel object model via COM interop; prefer `.values` when possible
5. **Single range input**: Each backtick expression takes one range reference as input

## Future Roadmap

See the [implementation plan](docs/specs%20and%20plan/excel-udf-addin-implementation-plan.md) for full details.

**Next phases:**
- Phase 8-9: Floating editor with syntax highlighting and autocomplete
- Phase 10-11: Named UDFs and persistence
- Phase 12: VBA transpiler for portable workbooks
- Phase 13: Built-in algorithm library (graph algorithms, combinatorics)

## Documentation

- [Full Specification](docs/specs%20and%20plan/excel-udf-addin-spec.md)
- [Implementation Plan](docs/specs%20and%20plan/excel-udf-addin-implementation-plan.md)
- [Technical Notes (CLAUDE.md)](CLAUDE.md) - includes important lessons about ExcelDNA assembly identity issues

## License

[TBD]
