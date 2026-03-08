using FormulaBoss.Interception;

namespace FormulaBoss.UI;

/// <summary>
///     Helper methods for completion scenarios not handled by Roslyn:
///     bracket column access (r["col"]) and non-DSL top-level completions.
/// </summary>
internal static class CompletionHelpers
{
    /// <summary>
    ///     Gets the length of the current partial word being typed (for replacement range).
    ///     In bracket context, includes spaces so the entire bracket content is replaced.
    /// </summary>
    public static int GetWordLength(string textUpToCaret, bool isBracketContext = false)
    {
        if (isBracketContext)
        {
            var i = textUpToCaret.Length - 1;
            while (i >= 0 && textUpToCaret[i] != '[')
            {
                i--;
            }

            return i >= 0 ? textUpToCaret.Length - 1 - i : 0;
        }

        var j = textUpToCaret.Length - 1;
        while (j >= 0 && char.IsLetterOrDigit(textUpToCaret[j]))
        {
            j--;
        }

        return textUpToCaret.Length - 1 - j;
    }

    /// <summary>
    ///     Returns column completions for dot access contexts (r.).
    ///     Displayed as dot notation but inserted as bracket syntax.
    /// </summary>
    public static IReadOnlyList<CompletionData> GetDotColumnCompletions(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata)
    {
        var ctx = ContextResolver.Resolve(textUpToCaret, metadata);
        if (ctx.Type != DslType.Row)
        {
            return Array.Empty<CompletionData>();
        }

        var tableName = ResolveTableName(ctx.TableName, fullText, metadata);
        return BuildRowCompletions(metadata, false, tableName);
    }

    /// <summary>
    ///     Returns column completions for bracket access contexts (r[).
    /// </summary>
    public static IReadOnlyList<CompletionData> GetBracketCompletions(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata)
    {
        var ctx = ContextResolver.Resolve(textUpToCaret, metadata);
        if (ctx.Type != DslType.Row)
        {
            return Array.Empty<CompletionData>();
        }

        var tableName = ResolveTableName(ctx.TableName, fullText, metadata);
        return BuildRowCompletions(metadata, true, tableName);
    }

    /// <summary>
    ///     Completions outside DSL backtick regions — table/range names and LET variables.
    /// </summary>
    public static IReadOnlyList<CompletionData> GetNonDslTopLevel(string fullText, WorkbookMetadata? metadata)
    {
        var items = new List<CompletionData>();

        var letBindingNames = ExtractLetBindingNames(fullText);
        var shadowedNames = new HashSet<string>(letBindingNames, StringComparer.OrdinalIgnoreCase);

        foreach (var name in letBindingNames)
        {
            items.Add(new CompletionData(name, "LET variable") { Priority = 1 });
        }

        if (metadata != null)
        {
            foreach (var table in metadata.TableNames)
            {
                if (!shadowedNames.Contains(table))
                {
                    items.Add(new CompletionData(table, "Table") { Priority = 1 });
                }
            }

            foreach (var name in metadata.NamedRanges)
            {
                if (!shadowedNames.Contains(name))
                {
                    items.Add(new CompletionData(name, "Named range") { Priority = 1 });
                }
            }
        }

        return items;
    }

    private static string? ResolveTableName(string? tableName, string fullText, WorkbookMetadata? metadata)
    {
        if (tableName == null || metadata == null)
        {
            return tableName;
        }

        if (metadata.TableColumns.ContainsKey(tableName))
        {
            return tableName;
        }

        var bindings = ExtractLetBindings(fullText);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var current = tableName;
        while (current != null && visited.Add(current))
        {
            if (metadata.TableColumns.ContainsKey(current))
            {
                return current;
            }

            bindings.TryGetValue(current, out current);
        }

        return tableName;
    }

    private static IReadOnlyList<CompletionData> BuildRowCompletions(
        WorkbookMetadata? metadata, bool isBracketContext, string? tableName)
    {
        var items = new List<CompletionData>();

        if (metadata != null)
        {
            if (tableName != null &&
                metadata.TableColumns.TryGetValue(tableName, out var tableSpecificCols))
            {
                foreach (var col in tableSpecificCols)
                {
                    items.Add(new ColumnCompletionData(col, "Column name", isBracketContext) { Priority = 1 });
                }
            }
            else
            {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cols in metadata.TableColumns.Values)
                {
                    foreach (var col in cols)
                    {
                        if (seen.Add(col))
                        {
                            items.Add(new ColumnCompletionData(col, "Column name", isBracketContext) { Priority = 1 });
                        }
                    }
                }
            }
        }

        return items;
    }

    private static Dictionary<string, string> ExtractLetBindings(string fullText)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (LetFormulaParser.TryParse(fullText, out var structure) && structure != null)
        {
            foreach (var binding in structure.Bindings)
            {
                if (!string.IsNullOrWhiteSpace(binding.VariableName))
                {
                    var value = binding.Value.Trim();
                    if (value.Length > 0 &&
                        value.All(c => char.IsLetterOrDigit(c) || c == '_'))
                    {
                        result[binding.VariableName] = value;
                    }
                }
            }

            return result;
        }

        var letIdx = fullText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return result;
        }

        var bodyStart = letIdx + 4;
        var args = LetArgumentSplitter.SplitTolerant(fullText, bodyStart);

        for (var i = 0; i + 1 < args.Count; i += 2)
        {
            if (args.Count % 2 == 1 && i == args.Count - 1)
            {
                break;
            }

            var name = args[i].Trim();
            var value = args[i + 1].Trim();
            if (name.Length > 0 && char.IsLetter(name[0]) && name.All(c => char.IsLetterOrDigit(c) || c == '_') &&
                value.Length > 0 && value.All(c => char.IsLetterOrDigit(c) || c == '_'))
            {
                result[name] = value;
            }
        }

        return result;
    }

    private static List<string> ExtractLetBindingNames(string fullText)
    {
        var names = new List<string>();

        if (LetFormulaParser.TryParse(fullText, out var structure) && structure != null)
        {
            foreach (var binding in structure.Bindings)
            {
                if (!string.IsNullOrWhiteSpace(binding.VariableName))
                {
                    names.Add(binding.VariableName);
                }
            }

            return names;
        }

        var letIdx = fullText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return names;
        }

        var bodyStart = letIdx + 4;
        var args = LetArgumentSplitter.SplitTolerant(fullText, bodyStart);

        for (var i = 0; i < args.Count; i += 2)
        {
            if (args.Count % 2 == 1 && i == args.Count - 1)
            {
                break;
            }

            var name = args[i].Trim();
            if (name.Length > 0 && char.IsLetter(name[0]) && name.All(c => char.IsLetterOrDigit(c) || c == '_'))
            {
                names.Add(name);
            }
        }

        return names;
    }

}
