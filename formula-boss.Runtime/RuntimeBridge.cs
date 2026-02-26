namespace FormulaBoss.Runtime;

/// <summary>
///     Static delegate fields for COM operations. The host (AddIn.AutoOpen) initialises these
///     with lambdas that perform actual COM/interop calls. Runtime types invoke delegates
///     without any direct COM dependency, avoiding assembly identity issues.
/// </summary>
public static class RuntimeBridge
{
    /// <summary>
    ///     Resolves a cell at the given absolute worksheet position.
    ///     Parameters: (sheetName, row1Based, col1Based) → Cell.
    ///     Initialised by the host at startup.
    /// </summary>
    public static Func<string, int, int, Cell>? GetCell { get; set; }

    /// <summary>
    ///     Resolves table header names from a range reference.
    ///     Parameters: (rangeRef) → string[]? (header names, or null if not a table).
    ///     Initialised by the host at startup.
    /// </summary>
    public static Func<object, string[]?>? GetHeaders { get; set; }

    /// <summary>
    ///     Resolves the origin (sheet name, top-left position) of a range reference.
    ///     Parameters: (rangeRef) → RangeOrigin? (or null if unavailable).
    ///     Initialised by the host at startup.
    /// </summary>
    public static Func<object, RangeOrigin?>? GetOrigin { get; set; }
}
