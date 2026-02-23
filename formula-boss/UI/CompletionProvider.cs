using FormulaBoss.Interception;
using FormulaBoss.Transpilation;

namespace FormulaBoss.UI;

/// <summary>
///     Provides context-aware completion items for the Formula Boss DSL.
///     Uses token-based backward walk to determine the type context at the caret.
/// </summary>
internal static class CompletionProvider
{
    private static readonly CompletionData[] MethodItems =
    {
        new("where", "Filter elements with a predicate") { Priority = 1 },
        new("select", "Map/transform elements") { Priority = 1 },
        new("map", "Transform preserving 2D shape") { Priority = 1 },
        new("find", "First element matching predicate") { Priority = 1 },
        new("some", "True if any element matches") { Priority = 1 },
        new("every", "True if all elements match") { Priority = 1 },
        new("orderBy", "Sort ascending by key") { Priority = 1 },
        new("orderByDesc", "Sort descending by key") { Priority = 1 },
        new("take", "First n elements") { Priority = 1 }, new("skip", "Skip first n elements") { Priority = 1 },
        new("distinct", "Remove duplicates") { Priority = 1 },
        new("groupBy", "Group elements by key") { Priority = 1 },
        new("reduce", "Fold to single value") { Priority = 1 },
        new("aggregate", "Fold to single value (alias)") { Priority = 1 },
        new("scan", "Running reduction") { Priority = 1 }, new("sum", "Sum of values") { Priority = 2 },
        new("avg", "Average of values") { Priority = 2 }, new("min", "Minimum value") { Priority = 2 },
        new("max", "Maximum value") { Priority = 2 }, new("count", "Count of elements") { Priority = 2 },
        new("first", "First element") { Priority = 2 },
        new("firstOrDefault", "First element or null") { Priority = 2 },
        new("last", "Last element") { Priority = 2 }, new("lastOrDefault", "Last element or null") { Priority = 2 },
        new("toArray", "Materialise to array") { Priority = 2 }
    };

    private static readonly CompletionData[] AccessorItems =
    {
        new("cells", "Iterate with object model access") { Priority = 1 },
        new("values", "Iterate values only (fast path)") { Priority = 1 },
        new("rows", "Iterate as Row objects") { Priority = 1 },
        new("cols", "Iterate columns as arrays") { Priority = 1 },
        new("withHeaders", "Treat first row as headers") { Priority = 2 }
    };

    private static readonly CompletionData[] KeywordItems =
    {
        new("var", "Variable declaration"), new("return", "Return a value"), new("if", "Conditional"),
        new("else", "Else branch"), new("true", "Boolean true"), new("false", "Boolean false"),
        new("null", "Null value"), new("new", "Object creation")
    };

    // DSL cell property aliases (the friendly names used in the DSL, not the raw COM names)
    private static readonly CompletionData[] CellPropertyItems =
    {
        new("value", "Cell value"), new("color", "Interior.ColorIndex"), new("rgb", "Interior.Color (RGB)"),
        new("bold", "Font.Bold"), new("italic", "Font.Italic"), new("fontSize", "Font.Size"),
        new("format", "NumberFormat"), new("formula", "Cell formula"), new("row", "Row number"),
        new("col", "Column number"), new("address", "Cell address"),
        new("Interior", "Interior object (ColorIndex, Color, Pattern)") { Priority = 2 },
        new("Font", "Font object (Bold, Italic, Size, Color)") { Priority = 2 }
    };

    /// <summary>
    ///     Returns completion items appropriate for the current cursor context.
    /// </summary>
    public static IReadOnlyList<CompletionData> GetCompletions(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata) =>
        GetCompletions(textUpToCaret, fullText, metadata, out _);

    /// <summary>
    ///     Returns completion items and whether the context is a bracket accessor (r[).
    /// </summary>
    public static IReadOnlyList<CompletionData> GetCompletions(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata, out bool isBracketContext)
    {
        var ctx = ContextResolver.Resolve(textUpToCaret, metadata);
        isBracketContext = ctx.IsBracketContext;

        // Outside DSL backticks — don't show DSL-specific completions
        if (!ctx.InsideDsl && !IsEntireFormulaBacktick(fullText))
        {
            return ctx.Type == DslType.TopLevel
                ? BuildNonDslTopLevel(fullText, metadata)
                : Array.Empty<CompletionData>();
        }

        return ctx.Type switch
        {
            DslType.Range => BuildRangeCompletions(),
            DslType.Pipeline => MethodItems,
            DslType.Cell => CellPropertyItems,
            DslType.Row => BuildRowCompletions(metadata, ctx.IsBracketContext, ResolveTableName(ctx.TableName, fullText, metadata)),
            DslType.Interior => BuildTypeProperties("Interior"),
            DslType.Font => BuildTypeProperties("Font"),
            DslType.Scalar => Array.Empty<CompletionData>(),
            DslType.TopLevel => BuildTopLevelCompletions(fullText, metadata),
            DslType.Unknown => BuildFallbackCompletions(),
            _ => Array.Empty<CompletionData>()
        };
    }

    /// <summary>
    ///     Gets the length of the current partial word being typed (for replacement range).
    ///     In bracket context, includes spaces so the entire bracket content is replaced.
    /// </summary>
    public static int GetWordLength(string textUpToCaret, bool isBracketContext = false)
    {
        if (isBracketContext)
        {
            // Walk back to the opening bracket, including spaces
            var i = textUpToCaret.Length - 1;
            while (i >= 0 && textUpToCaret[i] != '[')
            {
                i--;
            }

            // i is now at '[' or -1; word starts after '['
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
    ///     Checks if the entire formula is a backtick formula (starts with = and contains backticks),
    ///     or is a simple DSL expression without wrapping (no =LET structure).
    ///     In these cases the whole editor content is DSL context.
    /// </summary>
    private static bool IsEntireFormulaBacktick(string fullText)
    {
        // A formula that is just a backtick expression: =`expr` or '=`expr`
        // Or a simple non-LET formula containing backticks
        return BacktickExtractor.IsBacktickFormula(fullText);
    }

    /// <summary>
    ///     Range completions: show accessors (explicit path) + methods (implicit .values path).
    /// </summary>
    private static IReadOnlyList<CompletionData> BuildRangeCompletions()
    {
        var items = new List<CompletionData>(AccessorItems.Length + MethodItems.Length);
        items.AddRange(AccessorItems);
        items.AddRange(MethodItems);
        return items;
    }

    /// <summary>
    ///     Resolves a table name, following LET binding aliases if needed.
    ///     For example, if tableName is "t" and the formula has LET(t, Table1, ...),
    ///     returns "Table1".
    /// </summary>
    private static string? ResolveTableName(string? tableName, string fullText, WorkbookMetadata? metadata)
    {
        if (tableName == null || metadata == null)
        {
            return tableName;
        }

        // If the name directly matches a table, use it
        if (metadata.TableColumns.ContainsKey(tableName))
        {
            return tableName;
        }

        // Try to resolve as a LET binding
        var bindings = ExtractLetBindings(fullText);
        // Follow the chain (a LET var could alias another LET var)
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

        return tableName; // Couldn't resolve — return as-is (fallback to all columns)
    }

    /// <summary>
    ///     Extracts LET binding name→value pairs from the formula.
    ///     Values are trimmed identifiers only (not complex expressions).
    /// </summary>
    private static Dictionary<string, string> ExtractLetBindings(string fullText)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (LetFormulaParser.TryParse(fullText, out var structure) && structure != null)
        {
            foreach (var binding in structure.Bindings)
            {
                if (!string.IsNullOrWhiteSpace(binding.VariableName))
                {
                    var value = binding.Value?.Trim();
                    if (value != null && value.Length > 0 &&
                        value.All(c => char.IsLetterOrDigit(c) || c == '_'))
                    {
                        result[binding.VariableName] = value;
                    }
                }
            }

            return result;
        }

        // Fallback: tolerant extraction
        var letIdx = fullText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0) return result;

        var bodyStart = letIdx + 4;
        var args = SplitArgumentsTolerant(fullText, bodyStart);

        for (var i = 0; i + 1 < args.Count; i += 2)
        {
            if (args.Count % 2 == 1 && i == args.Count - 1) break;

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

    private static IReadOnlyList<CompletionData> BuildRowCompletions(
        WorkbookMetadata? metadata, bool isBracketContext, string? tableName)
    {
        var items = new List<CompletionData>();

        if (metadata != null)
        {
            // If we know the table, show only that table's columns
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
                // Fallback: show all columns across all tables (deduplicated)
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

    private static IReadOnlyList<CompletionData> BuildTypeProperties(string typeName)
    {
        if (!ExcelTypeSystem.Types.TryGetValue(typeName, out var props))
        {
            return Array.Empty<CompletionData>();
        }

        var items = new List<CompletionData>(props.Count);
        foreach (var prop in props.Values)
        {
            items.Add(new CompletionData(prop.Name, $"{prop.ResultType} — {typeName}.{prop.Name}"));
        }

        return items;
    }

    private static IReadOnlyList<CompletionData> BuildTopLevelCompletions(
        string fullText, WorkbookMetadata? metadata)
    {
        var items = new List<CompletionData>(KeywordItems);

        // Extract LET binding names that are in scope and add them
        var letBindingNames = ExtractLetBindingNames(fullText);
        var shadowedNames = new HashSet<string>(letBindingNames, StringComparer.OrdinalIgnoreCase);

        foreach (var name in letBindingNames)
        {
            items.Add(new CompletionData(name, "LET variable") { Priority = 1 });
        }

        if (metadata != null)
        {
            // Add table names (unless shadowed by a LET binding)
            foreach (var table in metadata.TableNames)
            {
                if (!shadowedNames.Contains(table))
                {
                    items.Add(new CompletionData(table, "Table") { Priority = 1 });
                }
            }

            // Add named ranges (unless shadowed by a LET binding)
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

    /// <summary>
    ///     Completions outside DSL backtick regions — just table/range names and LET variables.
    /// </summary>
    private static IReadOnlyList<CompletionData> BuildNonDslTopLevel(string fullText, WorkbookMetadata? metadata)
    {
        var items = new List<CompletionData>();

        // LET variables are valid in any calculation position, not just inside DSL
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

    /// <summary>
    ///     Extracts LET binding variable names from the full formula text.
    ///     Tolerant of incomplete formulas (missing closing paren, partial bindings)
    ///     since the user is actively editing in the editor.
    /// </summary>
    private static List<string> ExtractLetBindingNames(string fullText)
    {
        var names = new List<string>();

        // First try the strict parser (handles complete formulas correctly)
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

        // Fallback: tolerant extraction for incomplete LET formulas
        // Look for =LET( prefix, then extract comma-separated arguments
        var letIdx = fullText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return names;
        }

        var bodyStart = letIdx + 4;
        var args = SplitArgumentsTolerant(fullText, bodyStart);

        // LET args are pairs: name, value, name, value, ..., result
        // Every even-indexed arg (0, 2, 4, ...) is a variable name — except possibly the last
        // one if the count is odd (that's the result expression)
        for (var i = 0; i < args.Count; i += 2)
        {
            // Skip the last argument if odd count (it's the result expression)
            if (args.Count % 2 == 1 && i == args.Count - 1)
            {
                break;
            }

            var name = args[i].Trim();
            // Variable names should be simple identifiers
            if (name.Length > 0 && char.IsLetter(name[0]) && name.All(c => char.IsLetterOrDigit(c) || c == '_'))
            {
                names.Add(name);
            }
        }

        return names;
    }

    /// <summary>
    ///     Splits arguments from a LET body, tolerant of incomplete/unclosed formulas.
    ///     Respects parenthesis nesting, string literals, and backtick regions.
    /// </summary>
    private static List<string> SplitArgumentsTolerant(string text, int startPos)
    {
        var args = new List<string>();
        var current = new System.Text.StringBuilder();
        var depth = 0;
        var inString = false;
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

            if (inString)
            {
                current.Append(c);
                if (c == '"')
                {
                    // Check for doubled quote (escaped)
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        current.Append(text[i + 1]);
                        i++;
                    }
                    else
                    {
                        inString = false;
                    }
                }

                continue;
            }

            switch (c)
            {
                case '`':
                    inBacktick = true;
                    current.Append(c);
                    break;
                case '"':
                    inString = true;
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
                        // Closing paren of LET — we're done
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

        // Incomplete formula — add whatever we have so far
        if (current.Length > 0)
        {
            args.Add(current.ToString());
        }

        return args;
    }

    private static IReadOnlyList<CompletionData> BuildFallbackCompletions()
    {
        var items = new List<CompletionData>(AccessorItems.Length + MethodItems.Length + CellPropertyItems.Length);
        items.AddRange(AccessorItems);
        items.AddRange(MethodItems);
        items.AddRange(CellPropertyItems);
        return items;
    }
}
