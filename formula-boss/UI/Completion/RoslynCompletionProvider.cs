using System.Text.RegularExpressions;

using FormulaBoss.Transpilation;

using Microsoft.CodeAnalysis.Completion;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Provides code completions using Roslyn's <see cref="CompletionService" />.
///     Falls back to the legacy provider for bracket contexts and non-DSL regions.
/// </summary>
internal sealed class RoslynCompletionProvider
{
    private readonly RoslynWorkspaceManager _workspace;

    public RoslynCompletionProvider(RoslynWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    /// <summary>
    ///     Returns completion items for the current cursor context.
    /// </summary>
    public async Task<(IReadOnlyList<CompletionData> Items, bool IsBracketContext)> GetCompletionsAsync(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata,
        CancellationToken cancellationToken)
    {
        // Use the existing context resolver for bracket detection and non-DSL contexts
        var ctx = ContextResolver.Resolve(textUpToCaret, metadata);

        // Bracket context: use column completions (Roslyn can't help with string indexers)
        if (ctx.IsBracketContext)
        {
            var items = CompletionHelpers.GetBracketCompletions(textUpToCaret, fullText, metadata);
            return (items, true);
        }

        // Outside DSL: use top-level completions (table names, LET variables).
        // Check cursor context (InsideDsl), not whether formula contains backticks elsewhere.
        if (!ctx.InsideDsl)
        {
            if (ctx.Type == DslType.TopLevel)
            {
                var items = CompletionHelpers.GetNonDslTopLevel(fullText, metadata);
                return (items, false);
            }

            return (Array.Empty<CompletionData>(), false);
        }

        // Inside DSL: build synthetic document and get Roslyn completions
        var (syntheticSource, caretOffset) = SyntheticDocumentBuilder.Build(
            fullText, textUpToCaret, metadata);

        var roslynItems = await _workspace.GetCompletionsAsync(
            syntheticSource, caretOffset, cancellationToken);

        // Get the type before the dot once for both Row and Table detection.
        // We need this even when Roslyn returns no items (e.g. lambda parameter on Row)
        // so we can still emit column completions.
        var typeName = metadata != null
            ? await _workspace.GetTypeBeforeDotAsync(caretOffset, cancellationToken)
            : null;

        // Check if the expression before the dot is a Row type — if so, augment Roslyn's
        // completions (LINQ/ExcelArray surface) with bracket-inserting column completions.
        // Filter the synthetic column-name properties from Roslyn so they don't duplicate
        // the bracket-syntax versions emitted by BuildRowCompletions.
        // Columns are listed first because they're the most common Row completion target —
        // the AvalonEdit CompletionList renders items in insertion order.
        var rowTableName = ResolveTableNameFromType(typeName, metadata, @"^__(.+)Row$");
        if (rowTableName != null)
        {
            var sanitisedColumns = GetSanitisedColumnNames(rowTableName, metadata);
            var linqItems = MapCompletionItems(roslynItems);
            linqItems.RemoveAll(item => sanitisedColumns.Contains(item.Text));

            var rowResult = new List<CompletionData>();
            rowResult.AddRange(CompletionHelpers.BuildRowCompletions(metadata, false, rowTableName));
            rowResult.AddRange(linqItems);
            return (rowResult, false);
        }

        if (roslynItems.Count == 0)
        {
            return (Array.Empty<CompletionData>(), false);
        }

        // Check if the expression before the dot is a Table type — if so,
        // augment Roslyn's completions with bracket-inserting column completions
        var tableTableName = ResolveTableNameFromType(typeName, metadata, @"^__(.+)Table$");
        if (tableTableName != null)
        {
            var roslynResult = MapCompletionItems(roslynItems);
            var columnItems = CompletionHelpers.BuildRowCompletions(metadata, false, tableTableName);
            roslynResult.AddRange(columnItems);
            return (roslynResult, false);
        }

        var result = MapCompletionItems(roslynItems);
        return (result, false);
    }

    /// <summary>
    ///     Returns the set of sanitised column-name property identifiers emitted on the
    ///     synthetic Row class for a given table. Used to strip these synthetic stubs
    ///     from Roslyn's completion list so they don't duplicate the bracket-syntax
    ///     column completions.
    /// </summary>
    private static HashSet<string> GetSanitisedColumnNames(string tableName, WorkbookMetadata? metadata)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (metadata == null ||
            !metadata.TableColumns.TryGetValue(tableName, out var columns))
        {
            return result;
        }

        foreach (var (sanitised, _) in ColumnMapper.BuildMapping(columns.ToArray()))
        {
            result.Add(sanitised);
        }

        return result;
    }

    /// <summary>
    ///     Resolves a synthetic type name back to the original table name using the given regex pattern.
    ///     Returns null if the type doesn't match or the table name can't be resolved.
    /// </summary>
    private static string? ResolveTableNameFromType(
        string? typeName, WorkbookMetadata? metadata, string pattern)
    {
        if (typeName == null || metadata == null)
        {
            return null;
        }

        var coreName = typeName.EndsWith("?") ? typeName[..^1] : typeName;
        var match = Regex.Match(coreName, pattern);
        if (!match.Success)
        {
            return null;
        }

        var sanitisedName = match.Groups[1].Value;

        // Exclude RowCollection matches when looking for Row types
        if (sanitisedName.EndsWith("RowCollection"))
        {
            return null;
        }

        // Resolve sanitised name back to original table name
        foreach (var (originalName, _) in metadata.TableColumns)
        {
            if (ColumnMapper.Sanitise(originalName) == sanitisedName)
            {
                return originalName;
            }
        }

        return null;
    }

    private static List<CompletionData> MapCompletionItems(IReadOnlyList<CompletionItem> items)
    {
        var result = new List<CompletionData>();

        foreach (var item in items)
        {
            var text = item.DisplayText;

            // Skip synthetic internal types
            if (text.StartsWith("__", StringComparison.Ordinal))
            {
                continue;
            }

            // Skip common noise
            if (IsNoiseItem(text))
            {
                continue;
            }

            var priority = GetPriority(item);
            var description = GetDescription(item);

            result.Add(new CompletionData(text, description) { Priority = priority });
        }

        return result;
    }

    private static bool IsNoiseItem(string text) =>
        text is "Equals" or "GetHashCode" or "GetType" or "ToString" or "ReferenceEquals"
            or "MemberwiseClone" or "Finalize";

    private static double GetPriority(CompletionItem item)
    {
        // Boost members over keywords/types for DSL context
        var tags = item.Tags;
        if (tags.Contains("Property") || tags.Contains("Method"))
        {
            return 1;
        }

        if (tags.Contains("Keyword"))
        {
            return 0.5;
        }

        return 0;
    }

    private static string? GetDescription(CompletionItem item)
    {
        // Use inline description if available
        if (!string.IsNullOrEmpty(item.InlineDescription))
        {
            return item.InlineDescription;
        }

        // Use tags for a brief category
        if (item.Tags.Length > 0)
        {
            return item.Tags[0];
        }

        return null;
    }
}
