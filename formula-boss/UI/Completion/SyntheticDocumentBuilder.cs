using System.Text;

using FormulaBoss.Interception;
using FormulaBoss.Transpilation;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Builds a synthetic C# document for Roslyn <c>CompletionService</c>.
///     The document declares typed variables matching the user's LET bindings,
///     typed Row classes for each table's columns, and embeds the DSL expression
///     at a cursor position so Roslyn can provide completions.
/// </summary>
internal static class SyntheticDocumentBuilder
{
    /// <summary>
    ///     Builds a synthetic C# document and returns the source text plus the cursor offset
    ///     where Roslyn should provide completions.
    /// </summary>
    /// <param name="fullText">The full formula text (e.g. =LET(t, Table1, `t.rows.where(r => r.`))</param>
    /// <param name="textUpToCaret">The formula text up to the cursor position.</param>
    /// <param name="metadata">Workbook metadata with table/column info.</param>
    /// <returns>The synthetic source and the character offset for the cursor.</returns>
    public static (string Source, int CaretOffset) Build(
        string fullText, string textUpToCaret, WorkbookMetadata? metadata)
    {
        var sb = new StringBuilder(1024);

        // Usings
        AppendUsings(sb);

        // Generate typed row classes and typed table classes per table
        var tableTypeNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (metadata != null)
        {
            foreach (var (tableName, columns) in metadata.TableColumns)
            {
                var safeTableName = SanitiseIdentifier(tableName);
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
        }

        // Extract DSL expression from the formula
        var dslExpression = ExtractDslExpression(textUpToCaret);

        // Build the method body with variable declarations and the expression
        sb.AppendLine("class __Ctx {");
        sb.AppendLine("void __M() {");

        // Declare LET binding variables with appropriate types
        var bindings = ExtractLetBindingsTyped(fullText, metadata, tableTypeNames);
        var declaredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (varName, typeName) in bindings)
        {
            sb.AppendLine($"{typeName} {varName} = default!;");
            declaredNames.Add(varName);
        }

        // Declare table names and named ranges as variables (for direct use without LET)
        if (metadata != null)
        {
            foreach (var tableName in metadata.TableNames)
            {
                if (!declaredNames.Contains(tableName) && IsValidIdentifier(tableName) &&
                    tableTypeNames.TryGetValue(tableName, out var tableTypeName))
                {
                    sb.AppendLine($"{tableTypeName} {tableName} = default!;");
                    declaredNames.Add(tableName);
                }
            }

            foreach (var name in metadata.NamedRanges)
            {
                if (!declaredNames.Contains(name) && IsValidIdentifier(name))
                {
                    sb.AppendLine($"ExcelArray {name} = default!;");
                    declaredNames.Add(name);
                }
            }
        }

        // Embed the expression and mark cursor position
        if (IsStatementBlock(dslExpression))
        {
            // Statement block: emit as local function body so var/return/etc. are valid
            sb.Append("object __userBlock() ");
            var caretOffset = sb.Length + dslExpression.Length;
            sb.AppendLine(dslExpression);
            sb.AppendLine("var __result = __userBlock();");
            sb.AppendLine("}");
            sb.AppendLine("}");
            return (sb.ToString(), caretOffset);
        }
        else
        {
            sb.Append("var __result = ");
            var caretOffset = sb.Length + dslExpression.Length;
            sb.Append(dslExpression);
            sb.AppendLine(";");
            sb.AppendLine("}");
            sb.AppendLine("}");
            return (sb.ToString(), caretOffset);
        }
    }

    /// <summary>
    ///     Builds a synthetic C# document for Roslyn diagnostics and returns position mapping
    ///     so diagnostic spans can be translated back to editor offsets.
    /// </summary>
    public static DiagnosticBuildResult? BuildForDiagnostics(string fullText, WorkbookMetadata? metadata)
    {
        // Find the last backtick expression
        var expressions = BacktickExtractor.Extract(fullText);
        if (expressions.Count == 0)
        {
            return null;
        }

        var lastExpr = expressions[^1];
        var expressionText = lastExpr.Expression;
        // +1 to skip the opening backtick character
        var expressionStartInEditor = lastExpr.StartIndex + 1;

        var sb = new StringBuilder(1024);

        // Usings
        AppendUsings(sb);

        // Generate typed row classes and typed table classes per table
        var tableTypeNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (metadata != null)
        {
            foreach (var (tableName, columns) in metadata.TableColumns)
            {
                var safeTableName = SanitiseIdentifier(tableName);
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
        }

        sb.AppendLine("class __Ctx {");
        sb.AppendLine("void __M() {");

        // Declare LET binding variables with appropriate types
        var bindings = ExtractLetBindingsTyped(fullText, metadata, tableTypeNames);
        var declaredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (varName, typeName) in bindings)
        {
            sb.AppendLine($"{typeName} {varName} = default!;");
            declaredNames.Add(varName);
        }

        // Declare table names and named ranges as variables
        if (metadata != null)
        {
            foreach (var tableName in metadata.TableNames)
            {
                if (!declaredNames.Contains(tableName) && IsValidIdentifier(tableName) &&
                    tableTypeNames.TryGetValue(tableName, out var tableTypeName))
                {
                    sb.AppendLine($"{tableTypeName} {tableName} = default!;");
                    declaredNames.Add(tableName);
                }
            }

            foreach (var name in metadata.NamedRanges)
            {
                if (!declaredNames.Contains(name) && IsValidIdentifier(name))
                {
                    sb.AppendLine($"ExcelArray {name} = default!;");
                    declaredNames.Add(name);
                }
            }
        }

        // Embed the full expression and track its position
        int expressionStartInSynthetic;
        if (IsStatementBlock(expressionText))
        {
            sb.Append("object __userBlock() ");
            expressionStartInSynthetic = sb.Length;
            sb.AppendLine(expressionText);
            sb.AppendLine("var __result = __userBlock();");
        }
        else
        {
            sb.Append("var __result = ");
            expressionStartInSynthetic = sb.Length;
            sb.Append(expressionText);
            sb.AppendLine(";");
        }

        sb.AppendLine("}");
        sb.AppendLine("}");

        return new DiagnosticBuildResult(
            sb.ToString(),
            expressionStartInSynthetic,
            expressionStartInEditor,
            expressionText.Length);
    }

    private static void AppendUsings(StringBuilder sb)
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

    /// <summary>
    ///     Extracts the DSL expression from the text up to the cursor.
    ///     If inside backticks, extracts from the last unmatched backtick.
    ///     Otherwise treats the entire text as a potential expression.
    /// </summary>
    private static string ExtractDslExpression(string textUpToCaret)
    {
        // Find the last unmatched backtick
        var backtickCount = 0;
        var lastUnmatchedBacktick = -1;

        for (var i = 0; i < textUpToCaret.Length; i++)
        {
            if (textUpToCaret[i] == '`')
            {
                backtickCount++;
                if (backtickCount % 2 == 1)
                {
                    lastUnmatchedBacktick = i;
                }
            }
        }

        if (backtickCount % 2 == 1 && lastUnmatchedBacktick >= 0)
        {
            return textUpToCaret[(lastUnmatchedBacktick + 1)..];
        }

        // Not inside backticks — could be a simple backtick formula or just an expression
        // Try to find the last backtick pair and use content after it
        if (BacktickExtractor.IsBacktickFormula(textUpToCaret))
        {
            var lastBacktick = textUpToCaret.LastIndexOf('`');
            if (lastBacktick >= 0)
            {
                return textUpToCaret[(lastBacktick + 1)..];
            }
        }

        // Fallback: just return the text after any = sign
        var eqIdx = textUpToCaret.IndexOf('=');
        return eqIdx >= 0 ? textUpToCaret[(eqIdx + 1)..] : textUpToCaret;
    }

    /// <summary>
    ///     Extracts LET bindings with resolved type names.
    ///     Each binding value is checked against table names to assign the typed table class.
    /// </summary>
    private static List<(string VarName, string TypeName)> ExtractLetBindingsTyped(
        string fullText, WorkbookMetadata? metadata,
        Dictionary<string, string> tableTypeNames)
    {
        var result = new List<(string, string)>();

        // Try strict parser first
        if (LetFormulaParser.TryParse(fullText, out var structure) && structure != null)
        {
            foreach (var binding in structure.Bindings)
            {
                AddBinding(result, binding.VariableName, binding.Value, metadata, tableTypeNames);
            }

            return result;
        }

        // Fallback: tolerant extraction for incomplete LET formulas (user is typing)
        var letIdx = fullText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return result;
        }

        var args = SplitArgumentsTolerant(fullText, letIdx + 4);
        for (var i = 0; i + 1 < args.Count; i += 2)
        {
            if (args.Count % 2 == 1 && i == args.Count - 1)
            {
                break;
            }

            AddBinding(result, args[i], args[i + 1], metadata, tableTypeNames);
        }

        return result;
    }

    private static void AddBinding(
        List<(string VarName, string TypeName)> result,
        string rawName, string rawValue,
        WorkbookMetadata? metadata,
        Dictionary<string, string> tableTypeNames)
    {
        var varName = rawName.Trim();
        if (!IsValidIdentifier(varName))
        {
            return;
        }

        var value = rawValue.Trim();
        var typeName = ResolveBindingType(value, metadata, tableTypeNames);
        result.Add((varName, typeName));
    }

    /// <summary>
    ///     Splits LET arguments tolerantly, respecting parens, strings, and backtick regions.
    /// </summary>
    private static List<string> SplitArgumentsTolerant(string text, int startPos)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var depth = 0;
        var inBacktick = false;

        for (var i = startPos; i < text.Length; i++)
        {
            var c = text[i];

            if (inBacktick)
            {
                current.Append(c);
                if (c == '`')
                {
                    inBacktick = false;
                }

                continue;
            }

            switch (c)
            {
                case '`':
                    inBacktick = true;
                    current.Append(c);
                    break;
                case '(':
                    depth++;
                    current.Append(c);
                    break;
                case ')':
                    if (depth > 0)
                    {
                        depth--;
                        current.Append(c);
                    }
                    else
                    {
                        if (current.Length > 0)
                        {
                            args.Add(current.ToString());
                        }

                        return args;
                    }

                    break;
                case ',' when depth == 0:
                    args.Add(current.ToString());
                    current.Clear();
                    break;
                default:
                    current.Append(c);
                    break;
            }
        }

        if (current.Length > 0)
        {
            args.Add(current.ToString());
        }

        return args;
    }

    private static string ResolveBindingType(
        string value, WorkbookMetadata? metadata,
        Dictionary<string, string> tableTypeNames)
    {
        if (metadata == null)
        {
            return "ExcelValue";
        }

        // Check if value is a table name
        if (tableTypeNames.TryGetValue(value, out var tableTypeName))
        {
            return tableTypeName;
        }

        // Check if it's a named range
        foreach (var name in metadata.NamedRanges)
        {
            if (name.Equals(value, StringComparison.OrdinalIgnoreCase))
            {
                return "ExcelArray";
            }
        }

        return "ExcelValue";
    }

    private static string SanitiseIdentifier(string name)
    {
        return ColumnMapper.Sanitise(name);
    }

    /// <summary>
    ///     Detects whether a DSL expression is a statement block.
    ///     More lenient than <see cref="InputDetector.IsStatementBlock"/> —
    ///     does not require 'return' since the user may still be typing.
    /// </summary>
    private static bool IsStatementBlock(string expression) =>
        expression.TrimStart().StartsWith('{');

    private static bool IsValidIdentifier(string name) =>
        name.Length > 0 && (char.IsLetter(name[0]) || name[0] == '_') &&
        name.All(c => char.IsLetterOrDigit(c) || c == '_');
}

internal sealed record DiagnosticBuildResult(
    string Source,
    int ExpressionStartInSynthetic,
    int ExpressionStartInEditor,
    int ExpressionLength);
