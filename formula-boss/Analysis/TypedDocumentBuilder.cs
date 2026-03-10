using System.Text;

using FormulaBoss.Transpilation;
using FormulaBoss.UI;

namespace FormulaBoss.Analysis;

/// <summary>
///     Generates synthetic typed classes (Row, RowCollection, Table) and usings
///     for Roslyn compilation contexts. Shared between the editor (completions,
///     diagnostics, hover) and the pipeline (semantic analysis).
/// </summary>
internal static class TypedDocumentBuilder
{
    /// <summary>
    ///     Appends standard usings for the synthetic compilation context.
    /// </summary>
    public static void AppendUsings(StringBuilder sb)
    {
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using System.Text;");
        sb.AppendLine("using System.Text.RegularExpressions;");
        sb.AppendLine("using FormulaBoss.Runtime;");
        sb.AppendLine();
    }

    /// <summary>
    ///     Emits typed Row, RowCollection, and Table classes for each table in the metadata,
    ///     and returns a dictionary mapping table names to their synthetic type names.
    /// </summary>
    /// <returns>
    ///     A case-insensitive dictionary mapping original table names to synthetic table type names
    ///     (e.g. "Sales" → "__SalesTable").
    /// </returns>
    public static Dictionary<string, string> EmitTableTypes(
        StringBuilder sb, WorkbookMetadata? metadata)
    {
        var tableTypeNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (metadata == null)
        {
            return tableTypeNames;
        }

        foreach (var (tableName, columns) in metadata.TableColumns)
        {
            var safeTableName = ColumnMapper.Sanitise(tableName);
            if (string.IsNullOrEmpty(safeTableName))
            {
                continue;
            }

            var rowTypeName = $"__{safeTableName}Row";
            var rowCollTypeName = $"__{safeTableName}RowCollection";
            var tableTypeName = $"__{safeTableName}Table";
            tableTypeNames[tableName] = tableTypeName;

            EmitTypedRow(sb, rowTypeName, columns);
            EmitTypedRowCollection(sb, rowCollTypeName, rowTypeName);
            EmitTypedTable(sb, tableTypeName, rowCollTypeName);
        }

        return tableTypeNames;
    }

    private static void EmitTypedRow(StringBuilder sb, string rowTypeName, IReadOnlyList<string> columns)
    {
        sb.AppendLine($"class {rowTypeName} {{");

        // Indexer for bracket access
        sb.AppendLine($"public ColumnValue this[string columnName] => default!;");
        sb.AppendLine($"public ColumnValue this[int index] => default!;");
        sb.AppendLine($"public int ColumnCount => 0;");

        var mapping = ColumnMapper.BuildMapping(columns.ToArray());
        foreach (var (sanitised, _) in mapping)
        {
            sb.AppendLine($"public ColumnValue {sanitised} => default!;");
        }

        sb.AppendLine("}");
        sb.AppendLine();
    }

    private static void EmitTypedRowCollection(
        StringBuilder sb, string rowCollTypeName, string rowTypeName)
    {
        sb.AppendLine($"class {rowCollTypeName} : IEnumerable<{rowTypeName}> {{");
        sb.AppendLine($"public {rowCollTypeName} Where(Func<{rowTypeName}, bool> predicate) => this;");
        sb.AppendLine($"public IExcelRange Select(Func<{rowTypeName}, object> selector) => default!;");
        sb.AppendLine($"public bool Any(Func<{rowTypeName}, bool> predicate) => default;");
        sb.AppendLine($"public bool All(Func<{rowTypeName}, bool> predicate) => default;");
        sb.AppendLine($"public {rowTypeName} First(Func<{rowTypeName}, bool> predicate) => default!;");
        sb.AppendLine($"public {rowTypeName}? FirstOrDefault(Func<{rowTypeName}, bool> predicate) => default;");
        sb.AppendLine($"public {rowCollTypeName} OrderBy(Func<{rowTypeName}, object> keySelector) => this;");
        sb.AppendLine($"public {rowCollTypeName} OrderByDescending(Func<{rowTypeName}, object> keySelector) => this;");
        sb.AppendLine($"public int Count() => 0;");
        sb.AppendLine($"public {rowCollTypeName} Take(int count) => this;");
        sb.AppendLine($"public {rowCollTypeName} Skip(int count) => this;");
        sb.AppendLine($"public {rowCollTypeName} Distinct() => this;");
        sb.AppendLine($"public IExcelRange ToRange() => default!;");
        sb.AppendLine($"public dynamic Aggregate(dynamic seed, Func<dynamic, {rowTypeName}, dynamic> func) => default!;");
        sb.AppendLine($"public IExcelRange Scan(dynamic seed, Func<dynamic, {rowTypeName}, dynamic> func) => default!;");
        sb.AppendLine($"public GroupedRowCollection GroupBy(Func<{rowTypeName}, object> keySelector) => default!;");
        sb.AppendLine($"public IEnumerator<{rowTypeName}> GetEnumerator() => default!;");
        sb.AppendLine($"IEnumerator IEnumerable.GetEnumerator() => default!;");
        sb.AppendLine("}");
        sb.AppendLine();
    }

    private static void EmitTypedTable(
        StringBuilder sb, string tableTypeName, string rowCollTypeName)
    {
        sb.AppendLine($"class {tableTypeName} : ExcelTable {{");
        sb.AppendLine($"public new {rowCollTypeName} Rows => default!;");
        sb.AppendLine($"{tableTypeName}() : base(new object[0,0], System.Array.Empty<string>()) {{}}");
        sb.AppendLine("}");
        sb.AppendLine();
    }
}
