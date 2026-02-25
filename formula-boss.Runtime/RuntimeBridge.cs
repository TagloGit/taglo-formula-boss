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
}
