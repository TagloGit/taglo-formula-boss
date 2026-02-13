using System.Windows.Media;

namespace FormulaBoss.UI.Animation;

public record SpriteFrame(int[][] Grid, string Label, bool Shake, double DurationMultiplier = 1.0);

public static class SpritePalette
{
    public static readonly Dictionary<int, Color> Colors = new()
    {
        [0] = System.Windows.Media.Colors.Transparent,
        [1] = Color.FromRgb(0x1a, 0x1a, 0x2e),  // dark blue-black
        [4] = Color.FromRgb(0xff, 0xd7, 0x00),  // gold
        [5] = System.Windows.Media.Colors.White,
        [49] = Color.FromRgb(0x99, 0x1b, 0x1b), // dark red
        [52] = Color.FromRgb(0xc2, 0x41, 0x0c), // burnt orange
        [53] = Color.FromRgb(0xfb, 0x92, 0x3c), // orange
        [55] = Color.FromRgb(0xe8, 0x85, 0x3a), // light orange-brown
        [56] = Color.FromRgb(0xd9, 0x75, 0x26), // medium orange
        [57] = Color.FromRgb(0xb8, 0x5e, 0x1a), // dark orange
        [58] = Color.FromRgb(0x9a, 0x4e, 0x15), // brown
    };
}
