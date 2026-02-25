namespace FormulaBoss.Runtime;

/// <summary>
///     Wrapper for cell interior (background) formatting properties.
/// </summary>
public class Interior
{
    public int ColorIndex { get; init; }
    public int Color { get; init; }

    public Interior() { }

    public Interior(int colorIndex, int color)
    {
        ColorIndex = colorIndex;
        Color = color;
    }
}
