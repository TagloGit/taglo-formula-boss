namespace FormulaBoss.Runtime;

/// <summary>
///     Wrapper for a single Excel cell, providing access to value and formatting properties.
///     Populated via <see cref="RuntimeBridge.GetCell" /> delegate — no direct COM dependency.
/// </summary>
public class Cell
{
    public object? Value { get; init; }
    public string Formula { get; init; } = "";
    public string Format { get; init; } = "";
    public string Address { get; init; } = "";
    public int Row { get; init; }
    public int Col { get; init; }
    public Interior Interior { get; init; } = new();
    public CellFont Font { get; init; } = new();

    // Shorthand properties delegating to sub-objects
    public int Color => Interior.ColorIndex;
    public int Rgb => Interior.Color;
    public bool Bold => Font.Bold;
    public bool Italic => Font.Italic;
    public double FontSize => Font.Size;
}
