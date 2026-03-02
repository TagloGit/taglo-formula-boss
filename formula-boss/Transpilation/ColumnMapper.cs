using System.Text;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Builds a mapping from sanitised C# identifiers to original column names.
///     Used by <see cref="DotNotationRewriter"/> to convert dot access to bracket access.
/// </summary>
public static class ColumnMapper
{
    /// <summary>
    ///     Sanitises a column name to a valid C# identifier.
    ///     Removes spaces and special characters, preserves letters, digits, and underscores.
    /// </summary>
    public static string Sanitise(string columnName)
    {
        var sb = new StringBuilder(columnName.Length);

        foreach (var c in columnName)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
        }

        var result = sb.ToString();

        if (result.Length == 0)
        {
            return result;
        }

        // Prefix with underscore if starts with a digit
        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        return result;
    }

    /// <summary>
    ///     Builds a mapping from sanitised identifier → original column name.
    ///     Columns that conflict (two originals map to the same sanitised name) are excluded.
    /// </summary>
    public static Dictionary<string, string> BuildMapping(string[] headers)
    {
        // First pass: group by sanitised name to detect conflicts
        var groups = new Dictionary<string, List<string>>();

        foreach (var header in headers)
        {
            var sanitised = Sanitise(header);
            if (string.IsNullOrEmpty(sanitised))
            {
                continue;
            }

            if (!groups.TryGetValue(sanitised, out var list))
            {
                list = new List<string>();
                groups[sanitised] = list;
            }

            list.Add(header);
        }

        // Second pass: only include non-conflicting mappings
        var mapping = new Dictionary<string, string>();

        foreach (var (sanitised, originals) in groups)
        {
            if (originals.Count == 1)
            {
                mapping[sanitised] = originals[0];
            }
            // Conflicts (count > 1) are silently excluded — user must use bracket access
        }

        return mapping;
    }
}
