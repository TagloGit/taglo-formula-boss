namespace FormulaBoss.Runtime;

public static class CellExtensions
{
    public static double Sum(this IEnumerable<Cell> cells) =>
        cells.Aggregate(0.0, (acc, c) => acc + Convert.ToDouble(c.Value));

    public static double Average(this IEnumerable<Cell> cells)
    {
        var count = 0;
        var sum = 0.0;
        foreach (var c in cells)
        {
            sum += Convert.ToDouble(c.Value);
            count++;
        }

        return count == 0 ? 0.0 : sum / count;
    }

    public static double Min(this IEnumerable<Cell> cells) =>
        cells.Aggregate(double.MaxValue, (acc, c) => Math.Min(acc, Convert.ToDouble(c.Value)));

    public static double Max(this IEnumerable<Cell> cells) =>
        cells.Aggregate(double.MinValue, (acc, c) => Math.Max(acc, Convert.ToDouble(c.Value)));

    public static int Count(this IEnumerable<Cell> cells) =>
        Enumerable.Count(cells);
}
