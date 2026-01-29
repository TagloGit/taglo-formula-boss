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

Type a formula with backtick expressions. The formula must start with `'=` (quote prefix) to prevent Excel from evaluating it immediately:

```
'=SUM(`A1:A10.values.where(v => v > 0)`)
```

When you press Enter, Formula Boss:
1. Detects the backtick expression
2. Parses and compiles it to a UDF
3. Rewrites your formula to: `=SUM(__udf_abc123(A1:A10))`

## DSL Syntax

### Data Access

| Syntax | Description |
|--------|-------------|
| `range.values` | Iterate over cell values only (fast path) |
| `range.cells` | Iterate over cells with object model access (slower, but gives access to formatting) |

### Methods

| Method | Description |
|--------|-------------|
| `.where(predicate)` | Filter items matching predicate |
| `.select(transform)` | Transform each item |
| `.toArray()` | Materialize results as array |
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

```
# Sum positive values
'=SUM(`A1:A100.values.where(v => v > 0)`)

# Filter by cell color (yellow = index 6)
'=`A1:A100.cells.where(c => c.color == 6).select(c => c.value).toArray()`

# Get values from bold cells
'=`A1:A100.cells.where(c => c.bold).select(c => c.value).toArray()`

# Top 5 values
'=`A1:A100.values.orderByDesc(v => v).take(5).toArray()`

# Count cells matching condition
'=`A1:A100.values.where(v => v > 50 && v < 100).count()`
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

1. **Quote prefix required**: Formulas must start with `'=` to prevent Excel from evaluating before interception
2. **No IntelliSense**: The DSL editor is just Excel's formula bar - no autocomplete or syntax highlighting yet
3. **Object model is slow**: Using `.cells` accesses the Excel object model via COM interop; prefer `.values` when possible
4. **Single range input**: Each backtick expression takes one range reference as input

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
