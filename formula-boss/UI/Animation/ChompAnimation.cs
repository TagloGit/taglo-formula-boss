namespace FormulaBoss.UI.Animation;

public static class ChompAnimation
{
    // Shorthand palette indices matching the HTML source
    private const int _ = 0, D = 1, W = 5, R = 49, O = 52, L = 53, E = 55, J = 56, K = 57;

    private static readonly int[][] Base =
    [
        [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _], // 0  horns
        [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _], // 1  horn base
        [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _], // 2  top head
        [_, D, E, E, D, R, R, R, R, R, R, R, R, R, R, D, E, E, D, _], // 3  head
        [_, D, E, E, D, R, R, R, R, R, R, R, R, R, R, D, E, E, D, _], // 4  head
        [_, D, E, E, D, R, W, W, R, R, R, R, W, W, R, D, E, E, D, _], // 5  eyes white
        [_, D, L, L, R, R, W, D, R, R, R, R, D, W, R, R, L, L, D, _], // 6  eyes pupil
        [_, D, L, L, R, R, R, R, R, R, R, R, R, R, R, R, L, L, D, _], // 7  cheeks
        [_, D, L, L, R, R, R, R, R, O, O, R, R, R, R, R, L, L, D, _], // 8  nose
        [_, D, L, L, R, R, R, R, R, R, R, R, R, R, R, R, L, L, D, _], // 9  under nose
        [_, D, J, J, D, D, D, D, D, D, D, D, D, D, D, D, J, J, D, _], // 10 mouth border
        [_, D, J, J, D, W, D, W, D, W, D, W, D, W, W, D, J, J, D, _], // 11 upper teeth
        [_, D, J, J, D, R, R, R, O, O, O, R, R, R, R, D, J, J, D, _], // 12 mouth interior
        [_, D, K, K, D, D, W, D, W, D, W, D, W, D, D, D, K, K, D, _], // 13 lower teeth
        [_, D, K, K, K, R, R, R, R, R, R, R, R, R, R, K, K, K, D, _], // 14 chin
        [_, D, K, K, K, D, D, D, D, D, D, D, D, D, D, K, K, K, D, _], // 15 chin border
        [D, E, E, K, K, D, _, _, _, _, _, _, _, _, D, K, K, E, E, D], // 16 fist top
        [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D], // 17 fist mid
        [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D], // 18 fist bottom
        [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D] // 19 fist base
    ];

    public static List<SpriteFrame> BuildFrames()
    {
        var frames = new List<SpriteFrame>
        {
            // Frame 0: Idle
            new(Clone(Base), "Idle", false, 2.5), // Frame 1: Anticipation - eyes widen
            BuildAnticipation(), // Frame 2: Mouth opening
            BuildOpening(), // Frame 3: Wide open
            BuildWideOpen(),
            // Frame 4: Maximum gape
            BuildMaxGape(),
            // Frame 5: CHOMP! Jaws slam shut
            BuildChomp(),
            // Frame 6: Impact hold
            BuildCrunch(),
            // Frame 7: Releasing - satisfied look
            BuildSatisfied(),
            // Frame 8: Settling
            new(Clone(Base), "Settling", false),
            // Frame 9: Idle
            new(Clone(Base), "Idle", false, 2.5)
        };

        return frames;
    }

    private static SpriteFrame BuildAnticipation()
    {
        var f = Clone(Base);
        f[5] = [_, D, E, E, D, R, W, W, W, R, R, W, W, W, R, D, E, E, D, _];
        f[6] = [_, D, L, L, R, R, W, D, W, R, R, W, D, W, R, R, L, L, D, _];
        return new SpriteFrame(f, "Anticipation", false, 1.5);
    }

    private static SpriteFrame BuildOpening()
    {
        int[][] f =
        [
            [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _],
            [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _],
            [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _],
            [_, D, E, E, D, R, R, R, R, R, R, R, R, R, R, D, E, E, D, _],
            [_, D, E, E, D, R, W, W, W, R, R, W, W, W, R, D, E, E, D, _],
            [_, D, L, L, R, R, W, D, W, R, R, W, D, W, R, R, L, L, D, _],
            [_, D, L, L, R, R, R, R, R, O, O, R, R, R, R, R, L, L, D, _],
            [_, D, J, J, D, D, D, D, D, D, D, D, D, D, D, D, J, J, D, _],
            [_, D, J, J, D, W, D, W, D, W, D, W, D, W, W, D, J, J, D, _],
            [_, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _],
            [_, D, J, J, D, R, R, R, O, O, O, R, R, R, R, D, J, J, D, _],
            [_, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _],
            [_, D, K, K, D, D, W, D, W, D, W, D, W, D, D, D, K, K, D, _],
            [_, D, K, K, K, R, R, R, R, R, R, R, R, R, R, K, K, K, D, _],
            [_, D, K, K, K, D, D, D, D, D, D, D, D, D, D, K, K, K, D, _],
            [D, E, E, K, K, D, _, _, _, _, _, _, _, _, D, K, K, E, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D]
        ];
        return new SpriteFrame(f, "Opening", false);
    }

    private static SpriteFrame BuildWideOpen()
    {
        int[][] f =
        [
            [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _],
            [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _],
            [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _],
            [_, D, L, L, D, R, W, D, R, R, R, R, D, W, R, D, L, L, D, _],
            [_, D, L, L, D, D, D, D, D, D, D, D, D, D, D, D, L, L, D, _],
            [_, D, J, J, W, D, W, D, W, D, W, D, W, D, W, D, J, J, D, _],
            [_, D, J, J, D, W, D, W, D, W, D, W, D, W, D, W, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, O, O, O, O, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, O, O, O, O, O, O, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, K, K, D, W, D, W, D, W, D, W, D, W, D, W, K, K, D, _],
            [_, D, K, K, W, D, W, D, W, D, W, D, W, D, W, D, K, K, D, _],
            [_, D, K, K, D, D, D, D, D, D, D, D, D, D, D, D, K, K, D, _],
            [_, D, K, K, K, R, R, R, R, R, R, R, R, R, R, K, K, K, D, _],
            [D, E, E, K, K, D, D, D, D, D, D, D, D, D, D, K, K, E, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D]
        ];
        return new SpriteFrame(f, "Wide Open", false, 0.5);
    }

    private static SpriteFrame BuildMaxGape()
    {
        int[][] f =
        [
            [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _],
            [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _],
            [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _],
            [_, D, L, L, D, R, R, D, R, R, R, R, D, R, R, D, L, L, D, _],
            [_, D, J, J, D, D, D, D, D, D, D, D, D, D, D, D, J, J, D, _],
            [_, D, J, J, W, D, W, D, W, D, W, D, W, D, W, D, J, J, D, _],
            [_, D, J, J, D, W, D, W, D, W, D, W, D, W, D, W, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, O, O, O, O, O, O, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, O, O, O, O, O, O, O, O, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, J, J, R, R, R, R, R, R, R, R, R, R, R, R, J, J, D, _],
            [_, D, K, K, D, W, D, W, D, W, D, W, D, W, D, W, K, K, D, _],
            [_, D, K, K, W, D, W, D, W, D, W, D, W, D, W, D, K, K, D, _],
            [_, D, K, K, D, D, D, D, D, D, D, D, D, D, D, D, K, K, D, _],
            [D, E, E, K, K, D, D, D, D, D, D, D, D, D, D, K, K, E, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D]
        ];
        return new SpriteFrame(f, "MAX GAPE", false, 1.8);
    }

    private static SpriteFrame BuildChomp()
    {
        int[][] f =
        [
            [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _],
            [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _],
            [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _],
            [_, D, E, E, D, R, R, R, R, R, R, R, R, R, R, D, E, E, D, _],
            [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _],
            [_, D, L, L, R, R, R, R, R, O, O, R, R, R, R, R, L, L, D, _],
            [_, D, J, J, D, D, D, D, D, D, D, D, D, D, D, D, J, J, D, _],
            [_, D, J, J, W, W, W, W, W, W, W, W, W, W, W, W, J, J, D, _],
            [_, D, K, K, D, D, D, D, D, D, D, D, D, D, D, D, K, K, D, _],
            [_, D, K, K, K, R, R, R, R, R, R, R, R, R, R, K, K, K, D, _],
            [_, D, K, K, K, D, D, D, D, D, D, D, D, D, D, K, K, K, D, _],
            [D, E, E, K, K, D, _, _, _, _, _, _, _, _, D, K, K, E, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D]
        ];
        return new SpriteFrame(f, "CHOMP!", true, 0.5);
    }

    private static SpriteFrame BuildCrunch()
    {
        int[][] f =
        [
            [_, _, _, _, _, L, L, _, _, _, _, _, _, L, L, _, _, _, _, _],
            [_, _, _, _, _, O, O, D, D, D, D, D, D, O, O, _, _, _, _, _],
            [_, _, D, D, D, D, R, R, R, R, R, R, R, R, D, D, D, D, _, _],
            [_, D, E, E, D, R, R, R, R, R, R, R, R, R, R, D, E, E, D, _],
            [_, D, E, E, D, R, R, D, D, R, R, D, D, R, R, D, E, E, D, _],
            [_, D, L, L, R, R, R, R, R, O, O, R, R, R, R, R, L, L, D, _],
            [_, D, J, J, D, D, D, D, D, D, D, D, D, D, D, D, J, J, D, _],
            [_, D, J, J, W, D, W, D, W, D, W, D, W, D, W, D, J, J, D, _],
            [_, D, K, K, D, D, D, D, D, D, D, D, D, D, D, D, K, K, D, _],
            [_, D, K, K, K, R, R, R, R, R, R, R, R, R, R, K, K, K, D, _],
            [_, D, K, K, K, D, D, D, D, D, D, D, D, D, D, K, K, K, D, _],
            [D, E, E, K, K, D, _, _, _, _, _, _, _, _, D, K, K, E, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D],
            [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D]
        ];
        return new SpriteFrame(f, "CRUNCH!", true, 1.5);
    }

    private static SpriteFrame BuildSatisfied()
    {
        var f = Clone(Base);
        f[5] = [_, D, E, E, D, R, W, W, R, R, R, R, W, W, R, D, E, E, D, _];
        f[6] = [_, D, L, L, R, R, D, W, R, R, R, R, W, D, R, R, L, L, D, _];
        return new SpriteFrame(f, "Satisfied", false, 1.8);
    }

    private static int[][] Clone(int[][] grid) =>
        grid.Select(row => (int[])row.Clone()).ToArray();
}
