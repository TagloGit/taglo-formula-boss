namespace FormulaBoss.Transpilation;

/// <summary>
/// Defines Excel object model types and their properties for DSL validation.
/// </summary>
public static class ExcelTypeSystem
{
    /// <summary>
    /// Defines a property with its name, result type, and code generation template.
    /// Template uses {0} for the target expression.
    /// </summary>
    public record PropertyDef(string Name, string ResultType, string Template);

    /// <summary>
    /// Known types and their properties.
    /// </summary>
    public static readonly Dictionary<string, Dictionary<string, PropertyDef>> Types = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Cell"] = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Interior"] = new("Interior", "Interior", "{0}.Interior"),
            ["Font"] = new("Font", "Font", "{0}.Font"),
            ["Value"] = new("Value", "object", "{0}.Value"),
            ["Row"] = new("Row", "int", "{0}.Row"),
            ["Column"] = new("Column", "int", "{0}.Column"),
            ["Address"] = new("Address", "string", "{0}.Address"),
            ["Formula"] = new("Formula", "string", "(string)({0}.Formula ?? \"\")"),
            ["NumberFormat"] = new("NumberFormat", "string", "(string)({0}.NumberFormat ?? \"General\")"),
        },
        ["Interior"] = new(StringComparer.OrdinalIgnoreCase)
        {
            ["ColorIndex"] = new("ColorIndex", "int", "(int)({0}.ColorIndex ?? 0)"),
            ["Color"] = new("Color", "int", "(int)({0}.Color ?? 0)"),
            ["Pattern"] = new("Pattern", "int", "(int)({0}.Pattern ?? 0)"),
            ["PatternColorIndex"] = new("PatternColorIndex", "int", "(int)({0}.PatternColorIndex ?? 0)"),
        },
        ["Font"] = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Bold"] = new("Bold", "bool", "(bool)({0}.Bold ?? false)"),
            ["Italic"] = new("Italic", "bool", "(bool)({0}.Italic ?? false)"),
            ["Underline"] = new("Underline", "int", "(int)({0}.Underline ?? 0)"),
            ["Strikethrough"] = new("Strikethrough", "bool", "(bool)({0}.Strikethrough ?? false)"),
            ["Size"] = new("Size", "double", "(double)({0}.Size ?? 11)"),
            ["Color"] = new("Color", "int", "(int)({0}.Color ?? 0)"),
            ["ColorIndex"] = new("ColorIndex", "int", "(int)({0}.ColorIndex ?? 0)"),
            ["Name"] = new("Name", "string", "(string)({0}.Name ?? \"\")"),
        },
    };

    /// <summary>
    /// Finds a similar property name in a type using Levenshtein distance.
    /// </summary>
    /// <param name="typeName">The type to search in.</param>
    /// <param name="propertyName">The property name to find a match for.</param>
    /// <param name="maxDistance">Maximum edit distance for a suggestion (default 2).</param>
    /// <returns>The closest matching property name, or null if none found within maxDistance.</returns>
    public static string? FindSimilar(string typeName, string propertyName, int maxDistance = 2)
    {
        if (!Types.TryGetValue(typeName, out var props))
        {
            return null;
        }

        string? bestMatch = null;
        var bestDistance = maxDistance + 1;

        foreach (var prop in props.Keys)
        {
            var distance = LevenshteinDistance(propertyName.ToLowerInvariant(), prop.ToLowerInvariant());
            if (distance < bestDistance)
            {
                bestDistance = distance;
                bestMatch = prop;
            }
        }

        return bestDistance <= maxDistance ? bestMatch : null;
    }

    /// <summary>
    /// Calculates the Levenshtein distance between two strings.
    /// </summary>
    private static int LevenshteinDistance(string s1, string s2)
    {
        if (string.IsNullOrEmpty(s1))
        {
            return string.IsNullOrEmpty(s2) ? 0 : s2.Length;
        }

        if (string.IsNullOrEmpty(s2))
        {
            return s1.Length;
        }

        var d = new int[s1.Length + 1, s2.Length + 1];

        for (var i = 0; i <= s1.Length; i++)
        {
            d[i, 0] = i;
        }

        for (var j = 0; j <= s2.Length; j++)
        {
            d[0, j] = j;
        }

        for (var i = 1; i <= s1.Length; i++)
        {
            for (var j = 1; j <= s2.Length; j++)
            {
                var cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }

        return d[s1.Length, s2.Length];
    }
}
