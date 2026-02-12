namespace FormulaBoss.UI;

/// <summary>
///     Provides context-aware completion items for the Formula Boss DSL.
///     Currently uses dummy data sets to validate performance and UX.
/// </summary>
internal static class CompletionProvider
{
    private static readonly CompletionData[] Methods =
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
        new("scan", "Running reduction") { Priority = 1 }, new("sum", "Sum of values") { Priority = 1 },
        new("avg", "Average of values") { Priority = 1 }, new("min", "Minimum value") { Priority = 1 },
        new("max", "Maximum value") { Priority = 1 }, new("count", "Count of elements") { Priority = 1 },
        new("first", "First element") { Priority = 1 },
        new("firstOrDefault", "First element or null") { Priority = 1 },
        new("last", "Last element") { Priority = 1 }, new("lastOrDefault", "Last element or null") { Priority = 1 },
        new("toArray", "Materialise to array") { Priority = 1 }
    };

    private static readonly CompletionData[] Accessors =
    {
        new("cells", "Iterate with object model access") { Priority = 2 },
        new("values", "Iterate values only (fast path)") { Priority = 2 },
        new("rows", "Iterate as Row objects") { Priority = 2 },
        new("cols", "Iterate columns as arrays") { Priority = 2 },
        new("withHeaders", "Treat first row as headers") { Priority = 2 }
    };

    private static readonly CompletionData[] CellProperties =
    {
        new("value", "Cell value"), new("color", "Interior.ColorIndex"), new("rgb", "Interior.Color (RGB)"),
        new("bold", "Font.Bold"), new("italic", "Font.Italic"), new("fontSize", "Font.Size"),
        new("format", "NumberFormat"), new("formula", "Cell formula"), new("row", "Row number"),
        new("col", "Column number"), new("address", "Cell address")
    };

    private static readonly CompletionData[] Keywords =
    {
        new("var", "Variable declaration"), new("return", "Return a value"), new("if", "Conditional"),
        new("else", "Else branch"), new("true", "Boolean true"), new("false", "Boolean false"),
        new("null", "Null value"), new("new", "Object creation")
    };

    /// <summary>
    ///     Returns completion items appropriate for the current cursor context.
    /// </summary>
    public static IReadOnlyList<CompletionData> GetCompletions(string textUpToCaret)
    {
        // After a dot -> show methods, accessors, and cell properties
        var lastDot = textUpToCaret.LastIndexOf('.');
        if (lastDot >= 0)
        {
            var afterDot = textUpToCaret[(lastDot + 1)..];
            if (afterDot.Length == 0 || afterDot.All(char.IsLetterOrDigit))
            {
                var items = new List<CompletionData>(Methods.Length + Accessors.Length + CellProperties.Length);
                items.AddRange(Accessors);
                items.AddRange(Methods);
                items.AddRange(CellProperties);
                return items;
            }
        }

        // Default: show keywords
        return Keywords;
    }

    /// <summary>
    ///     Gets the length of the current partial word being typed (for replacement range).
    /// </summary>
    public static int GetWordLength(string textUpToCaret)
    {
        var i = textUpToCaret.Length - 1;
        while (i >= 0 && char.IsLetterOrDigit(textUpToCaret[i]))
        {
            i--;
        }

        return textUpToCaret.Length - 1 - i;
    }
}
