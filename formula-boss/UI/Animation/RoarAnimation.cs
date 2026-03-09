namespace FormulaBoss.UI.Animation;

public static class RoarAnimation
{
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
            new(Clone(Base), "Idle", false, 2.0),
            // Frame 1: Tense — brow furrow
            BuildTense(),
            // Frame 2: Mouth opening — 1 extra interior row (21 rows)
            BuildOpening(),
            // Frame 3: Full ROAR — 2 extra interior rows (22 rows), shake
            BuildRoar(),
            // Frame 4: Hold roar — pupils shift left, shake
            BuildRoarHoldLeft(),
            // Frame 5: Hold roar — pupils shift right, shake
            BuildRoarHoldRight(),
            // Frame 6: Mouth closing — 1 extra row
            BuildClosing(),
            // Frame 7: Back to idle
            new(Clone(Base), "Idle", false, 2.0)
        };

        return frames;
    }

    private static SpriteFrame BuildTense()
    {
        var f = Clone(Base);
        // Brow furrow: dark pixel above each eye
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        return new SpriteFrame(f, "Tense", false);
    }

    private static SpriteFrame BuildOpening()
    {
        var f = Clone(Base);
        // Brow furrow
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        // Insert extra dark mouth interior row between upper teeth (11) and mouth interior (12)
        var extra = new[] { _, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _ };
        var result = new int[f.Length + 1][];
        Array.Copy(f, 0, result, 0, 12);
        result[12] = extra;
        Array.Copy(f, 12, result, 13, f.Length - 12);
        return new SpriteFrame(result, "Opening", false);
    }

    private static SpriteFrame BuildRoar()
    {
        var f = Clone(Base);
        // Angry brow
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        // Wide mouth: replace single interior row with 3 rows
        var mouthDark1 = new[] { _, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _ };
        var mouthDark2 = new[] { _, D, J, J, D, R, R, R, O, O, O, R, R, R, R, D, J, J, D, _ }; // tongue
        // Build: rows 0-11 (upper teeth), 3 mouth rows, rows 13+ (lower teeth onward)
        var result = new int[f.Length + 2][];
        Array.Copy(f, 0, result, 0, 12);
        result[12] = mouthDark1;
        result[13] = mouthDark2;
        result[14] = mouthDark1;
        Array.Copy(f, 13, result, 15, f.Length - 13);
        return new SpriteFrame(result, "ROAR!", true, 1.5);
    }

    private static SpriteFrame BuildRoarHoldLeft()
    {
        var f = Clone(Base);
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        // Pupils shift left
        f[6] = [_, D, L, L, R, D, W, _, R, R, R, D, W, _, R, R, L, L, D, _];
        // Wide mouth
        var mouthDark1 = new[] { _, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _ };
        var mouthDark2 = new[] { _, D, J, J, D, R, R, R, O, O, O, R, R, R, R, D, J, J, D, _ };
        var result = new int[f.Length + 2][];
        Array.Copy(f, 0, result, 0, 12);
        result[12] = mouthDark1;
        result[13] = mouthDark2;
        result[14] = mouthDark1;
        Array.Copy(f, 13, result, 15, f.Length - 13);
        return new SpriteFrame(result, "ROAR! (hold)", true, 1.5);
    }

    private static SpriteFrame BuildRoarHoldRight()
    {
        var f = Clone(Base);
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        // Pupils shift right
        f[6] = [_, D, L, L, R, R, _, W, D, R, R, R, _, W, D, R, L, L, D, _];
        // Wide mouth
        var mouthDark1 = new[] { _, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _ };
        var mouthDark2 = new[] { _, D, J, J, D, R, R, R, O, O, O, R, R, R, R, D, J, J, D, _ };
        var result = new int[f.Length + 2][];
        Array.Copy(f, 0, result, 0, 12);
        result[12] = mouthDark1;
        result[13] = mouthDark2;
        result[14] = mouthDark1;
        Array.Copy(f, 13, result, 15, f.Length - 13);
        return new SpriteFrame(result, "ROAR! (hold)", true, 1.5);
    }

    private static SpriteFrame BuildClosing()
    {
        var f = Clone(Base);
        f[4] = [_, D, E, E, D, R, R, D, R, R, R, R, D, R, R, D, E, E, D, _];
        var extra = new[] { _, D, J, J, D, R, R, R, R, R, R, R, R, R, R, D, J, J, D, _ };
        var result = new int[f.Length + 1][];
        Array.Copy(f, 0, result, 0, 12);
        result[12] = extra;
        Array.Copy(f, 12, result, 13, f.Length - 12);
        return new SpriteFrame(result, "Closing", false);
    }

    private static int[][] Clone(int[][] grid) =>
        grid.Select(row => (int[])row.Clone()).ToArray();
}
