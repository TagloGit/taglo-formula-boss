using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class CellExtensionsTests
{
    private static Cell MakeCell(double value, int color = 0) => new()
    {
        Value = value,
        Interior = new Interior(color, 0)
    };

    private static List<Cell> TestCells() => new()
    {
        MakeCell(10, color: 3),
        MakeCell(20, color: 6),
        MakeCell(30, color: 3),
        MakeCell(40, color: 6)
    };

    [Fact]
    public void Sum_AllCells() =>
        Assert.Equal(100.0, TestCells().Sum());

    [Fact]
    public void Sum_FilteredByColor() =>
        Assert.Equal(40.0, TestCells().Where(c => c.Color == 3).Sum());

    [Fact]
    public void Average_AllCells() =>
        Assert.Equal(25.0, TestCells().Average());

    [Fact]
    public void Min_AllCells() =>
        Assert.Equal(10.0, TestCells().Min());

    [Fact]
    public void Max_AllCells() =>
        Assert.Equal(40.0, TestCells().Max());

    [Fact]
    public void Count_FilteredByColor() =>
        Assert.Equal(2, TestCells().Where(c => c.Color == 6).Count());

    [Fact]
    public void Sum_EmptyCells() =>
        Assert.Equal(0.0, new List<Cell>().Sum());

    [Fact]
    public void Average_EmptyCells() =>
        Assert.Equal(0.0, new List<Cell>().Average());
}
