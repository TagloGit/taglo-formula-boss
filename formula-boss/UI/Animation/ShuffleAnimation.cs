namespace FormulaBoss.UI.Animation;

public static class ShuffleAnimation
{
    private const int _ = 0, D = 1, W = 5, R = 49, O = 52, L = 53, E = 55, J = 56, K = 57;

    private const int CanvasWidth = 24;
    private const int Pad = 2;

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
        [_, D, K, K, K, D, D, D, D, D, D, D, D, D, D, K, K, K, D, _] // 15 chin border
    ];

    // Fist rows (separate from body for independent vertical placement)
    private static readonly int[][] FistRows =
    [
        [D, E, E, K, K, D, _, _, _, _, _, _, _, _, D, K, K, E, E, D], // top
        [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D], // mid
        [D, E, L, D, K, K, D, _, _, _, _, _, _, D, K, K, D, L, E, D], // bottom
        [D, D, D, D, D, D, D, _, _, _, _, _, _, D, D, D, D, D, D, D] // base
    ];

    public static List<SpriteFrame> BuildFrames()
    {
        return
        [
            // Frame 0: Idle (centered)
            BuildFrame(0, false, false, 0, false, "Idle", 2.0),
            // Frame 1: Lean left — shift 1px left, right fist up, eyes left, squash
            BuildFrame(-1, false, true, -1, true, "Lean Left"),
            // Frame 2: Step left — shift 2px left, left fist up, eyes left
            BuildFrame(-2, true, false, -1, false, "Step Left"),
            // Frame 3: Return center — bounce
            BuildFrame(0, false, false, 0, true, "Bounce", 0.7),
            // Frame 4: Lean right — shift 1px right, left fist up, eyes right, squash
            BuildFrame(1, true, false, 1, true, "Lean Right"),
            // Frame 5: Step right — shift 2px right, right fist up, eyes right
            BuildFrame(2, false, true, 1, false, "Step Right"),
            // Frame 6: Return center — bounce
            BuildFrame(0, false, false, 0, true, "Bounce", 0.7),
            // Frame 7: Idle
            BuildFrame(0, false, false, 0, false, "Idle", 2.0)
        ];
    }

    private static SpriteFrame BuildFrame(
        int bodyOffset, bool leftFistUp, bool rightFistUp, int eyeShift, bool squash,
        string label, double durationMultiplier = 1.0)
    {
        var bodyRows = Clone(Base);

        // Apply eye shift
        if (eyeShift == -1)
        {
            // Pupils look left (swap W and D within each eye — pupil moves left)
            bodyRows[6] = [_, D, L, L, R, R, D, W, R, R, R, R, W, D, R, R, L, L, D, _];
        }
        else if (eyeShift == 1)
        {
            // Pupils look right
            bodyRows[6] = [_, D, L, L, R, R, R, W, D, R, R, R, R, W, D, R, L, L, D, _];
        }

        // Squash: remove duplicate head row 4 for bounce effect
        var bodyList = bodyRows.ToList();
        if (squash)
        {
            bodyList.RemoveAt(4);
        }

        // Shift all body rows into wider canvas
        var rows = new List<int[]>();
        foreach (var row in bodyList)
        {
            rows.Add(ShiftRow(row, bodyOffset));
        }

        var bodyEnd = rows.Count;

        // Left fist: cols 0-6, Right fist: cols 13-19 from original FistRows
        var leftFist = FistRows.Select(r => r[..7]).ToArray();
        var rightFist = FistRows.Select(r => r[13..20]).ToArray();

        var leftStart = leftFistUp ? -1 : 0;
        var rightStart = rightFistUp ? -1 : 0;

        var leftFistEnd = bodyEnd + leftStart + 4;
        var rightFistEnd = bodyEnd + rightStart + 4;
        var totalRows = Math.Max(leftFistEnd, rightFistEnd);

        while (rows.Count < totalRows)
        {
            rows.Add(new int[CanvasWidth]);
        }

        // Place left fist
        for (var fi = 0; fi < 4; fi++)
        {
            var rowIdx = bodyEnd + leftStart + fi;
            if (rowIdx < 0 || rowIdx >= totalRows)
            {
                continue;
            }

            var pixels = leftFist[fi];
            var startX = Pad + bodyOffset;
            for (var px = 0; px < pixels.Length; px++)
            {
                var x = startX + px;
                if (x >= 0 && x < CanvasWidth && pixels[px] != _)
                {
                    rows[rowIdx][x] = pixels[px];
                }
            }
        }

        // Place right fist
        for (var fi = 0; fi < 4; fi++)
        {
            var rowIdx = bodyEnd + rightStart + fi;
            if (rowIdx < 0 || rowIdx >= totalRows)
            {
                continue;
            }

            var pixels = rightFist[fi];
            var startX = Pad + bodyOffset + 13;
            for (var px = 0; px < pixels.Length; px++)
            {
                var x = startX + px;
                if (x >= 0 && x < CanvasWidth && pixels[px] != _)
                {
                    rows[rowIdx][x] = pixels[px];
                }
            }
        }

        return new SpriteFrame(rows.ToArray(), label, false, durationMultiplier);
    }

    private static int[] ShiftRow(int[] row, int offset)
    {
        var result = new int[CanvasWidth];
        var startX = Pad + offset;
        for (var i = 0; i < row.Length; i++)
        {
            var x = startX + i;
            if (x >= 0 && x < CanvasWidth)
            {
                result[x] = row[i];
            }
        }

        return result;
    }

    private static int[][] Clone(int[][] grid) =>
        grid.Select(row => (int[])row.Clone()).ToArray();
}
