namespace FormulaBoss.Runtime;

/// <summary>
///     Positional context for a range within a worksheet.
///     Coordinates are 1-based (matching Excel's row/column numbering).
/// </summary>
public record RangeOrigin(string SheetName, int TopRow, int LeftCol);
