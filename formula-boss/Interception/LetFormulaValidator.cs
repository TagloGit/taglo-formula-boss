namespace FormulaBoss.Interception;

/// <summary>
///     Represents a structural error in a LET formula with its position in the formula text.
/// </summary>
/// <param name="StartOffset">Start position in the formula text.</param>
/// <param name="Length">Character length of the error region.</param>
/// <param name="Message">Human-readable error description.</param>
public record LetError(int StartOffset, int Length, string Message);

/// <summary>
///     Validates the structure of LET formulas and returns errors with position information
///     suitable for displaying squiggly underlines in the editor.
///     Only reports errors when the formula appears complete (has a matching closing parenthesis)
///     to avoid false positives while the user is typing.
/// </summary>
public static class LetFormulaValidator
{
    /// <summary>
    ///     Validates a LET formula and returns any structural errors found.
    ///     Returns an empty list for non-LET formulas or incomplete formulas (user still typing).
    /// </summary>
    public static List<LetError> Validate(string formulaText)
    {
        var errors = new List<LetError>();

        if (!LetFormulaParser.IsLetFormula(formulaText))
        {
            return errors;
        }

        var letIdx = formulaText.IndexOf("LET(", StringComparison.OrdinalIgnoreCase);
        if (letIdx < 0)
        {
            return errors;
        }

        var openParenIdx = letIdx + 3;
        var closeParenIdx = LetArgumentSplitter.FindMatchingCloseParen(formulaText, openParenIdx);

        if (closeParenIdx == -1)
        {
            // Unbalanced parentheses — formula is incomplete.
            // Only report if the formula has content after LET( suggesting the user
            // intended to close it (e.g., has a trailing ) that doesn't match).
            var hasTrailingCloseParen = formulaText.IndexOf(')', openParenIdx + 1) >= 0;
            if (hasTrailingCloseParen)
            {
                errors.Add(new LetError(openParenIdx, 1,
                    "Unbalanced parentheses in LET formula"));
            }

            // Don't validate further — formula is structurally incomplete
            return errors;
        }

        // Formula has matching parens — validate argument structure
        var bodyStart = openParenIdx + 1;
        var args = LetArgumentSplitter.SplitWithPositions(formulaText, bodyStart);

        if (args.Count == 0)
        {
            errors.Add(new LetError(letIdx, closeParenIdx - letIdx + 1,
                "LET formula is empty — requires at least one binding and a result expression"));
            return errors;
        }

        if (args.Count == 1)
        {
            errors.Add(new LetError(letIdx, closeParenIdx - letIdx + 1,
                "LET formula requires at least one binding (name, value) and a result expression"));
            return errors;
        }

        if (args.Count % 2 == 0)
        {
            // Even number of args — missing result expression (e.g., name, value but no result)
            var lastArg = args[^1];
            errors.Add(new LetError(lastArg.StartOffset, lastArg.Length,
                "Missing result expression — LET requires name/value pairs followed by a result"));
        }

        // Validate variable names (every other argument starting from index 0)
        var maxNameIndex = args.Count % 2 == 0 ? args.Count : args.Count - 1;
        for (var i = 0; i < maxNameIndex; i += 2)
        {
            var nameArg = args[i];
            var name = nameArg.Value.Trim();

            if (string.IsNullOrWhiteSpace(name))
            {
                errors.Add(new LetError(nameArg.StartOffset, Math.Max(nameArg.Length, 1),
                    "Empty variable name"));
                continue;
            }

            var trimStart = nameArg.Value.Length - nameArg.Value.TrimStart().Length;
            var trimmedLength = name.Length;

            if (!IsValidLetVariableName(name))
            {
                errors.Add(new LetError(nameArg.StartOffset + trimStart, trimmedLength,
                    $"Invalid variable name '{name}'"));
            }
            else if (LooksLikeCellAddress(name))
            {
                errors.Add(new LetError(nameArg.StartOffset + trimStart, trimmedLength,
                    $"Variable name '{name}' conflicts with a cell address"));
            }
            else if (name.Length >= 255)
            {
                errors.Add(new LetError(nameArg.StartOffset + trimStart, trimmedLength,
                    "Variable name exceeds 255 characters"));
            }
        }

        return errors;
    }

    private static bool IsValidLetVariableName(string name) =>
        name.Length > 0 &&
        (char.IsLetter(name[0]) || name[0] == '_') &&
        name.All(c => char.IsLetterOrDigit(c) || c == '_');

    /// <summary>
    ///     Checks whether a name matches the pattern of an Excel cell address (1-3 letters + 1-7 digits).
    ///     Excel columns run A–XFD (max 3 letters) and rows 1–1048576 (max 7 digits).
    ///     The check is conservative: any letter prefix followed by digits is flagged,
    ///     even if the column is beyond XFD, since Excel still rejects these as variable names.
    /// </summary>
    private static bool LooksLikeCellAddress(string name)
    {
        // Must start with a letter (underscore-prefixed names are always safe)
        if (name[0] == '_')
        {
            return false;
        }

        // Find where letters end and digits begin
        var i = 0;
        while (i < name.Length && char.IsLetter(name[i]))
        {
            i++;
        }

        // Must have at least one letter and at least one digit, and nothing else
        if (i == 0 || i >= name.Length)
        {
            return false;
        }

        // Everything after the letters must be digits
        for (var j = i; j < name.Length; j++)
        {
            if (!char.IsDigit(name[j]))
            {
                return false;
            }
        }

        // Letter prefix 1-3 chars + digit suffix 1-7 chars matches cell address pattern
        var letterCount = i;
        var digitCount = name.Length - i;
        return letterCount <= 3 && digitCount is >= 1 and <= 7;
    }
}
