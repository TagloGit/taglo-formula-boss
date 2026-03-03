using FormulaBoss.Interception;

using Microsoft.CodeAnalysis.Completion;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Provides code completions using Roslyn's <see cref="CompletionService" />.
///     Falls back to the legacy provider for bracket contexts and non-DSL regions.
/// </summary>
internal sealed class RoslynCompletionProvider
{
    private readonly RoslynWorkspaceManager _workspace;

    /// <summary>
    ///     Synthetic type name prefixes to filter from Roslyn results.
    /// </summary>
    private static readonly HashSet<string> InternalPrefixes = new() { "__" };

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

        // Row dot context: show column names that insert as bracket syntax
        if (ctx.Type == DslType.Row)
        {
            var items = CompletionHelpers.GetDotColumnCompletions(textUpToCaret, fullText, metadata);
            return (items, false);
        }

        // Outside DSL: use legacy top-level completions (table names, LET variables)
        if (!ctx.InsideDsl && !BacktickExtractor.IsBacktickFormula(fullText))
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

        if (roslynItems.Count == 0)
        {
            return (Array.Empty<CompletionData>(), false);
        }

        var result = MapCompletionItems(roslynItems);
        return (result, false);
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
