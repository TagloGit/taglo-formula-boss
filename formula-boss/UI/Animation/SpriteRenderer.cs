using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FormulaBoss.UI.Animation;

public static class SpriteRenderer
{
    /// <summary>
    ///     Renders a sprite grid to a WriteableBitmap. Each grid cell becomes one pixel.
    ///     Scale up with NearestNeighbor in the Image control for crisp pixel art.
    /// </summary>
    public static WriteableBitmap Render(int[][] grid)
    {
        var rows = grid.Length;
        var cols = grid[0].Length;
        var bmp = new WriteableBitmap(cols, rows, 96, 96, PixelFormats.Bgra32, null);
        var pixels = new byte[rows * cols * 4];

        for (var y = 0; y < rows; y++)
        {
            for (var x = 0; x < cols; x++)
            {
                var paletteIdx = grid[y][x];
                if (!SpritePalette.Colors.TryGetValue(paletteIdx, out var color))
                {
                    continue;
                }

                var offset = ((y * cols) + x) * 4;
                pixels[offset + 0] = color.B;
                pixels[offset + 1] = color.G;
                pixels[offset + 2] = color.R;
                pixels[offset + 3] = paletteIdx == 0 ? (byte)0 : (byte)255;
            }
        }

        bmp.WritePixels(new Int32Rect(0, 0, cols, rows), pixels, cols * 4, 0);
        bmp.Freeze();
        return bmp;
    }
}
