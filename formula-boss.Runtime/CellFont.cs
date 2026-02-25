namespace FormulaBoss.Runtime;

/// <summary>
///     Wrapper for cell font formatting properties.
/// </summary>
public class CellFont
{
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public double Size { get; init; }
    public string Name { get; init; } = "";
    public int Color { get; init; }

    public CellFont() { }

    public CellFont(bool bold, bool italic, double size, string name, int color)
    {
        Bold = bold;
        Italic = italic;
        Size = size;
        Name = name;
        Color = color;
    }
}
