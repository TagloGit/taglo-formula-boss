using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class CellTests
{
    [Fact]
    public void Cell_ShorthandProperties_DelegateToSubObjects()
    {
        var cell = new Cell
        {
            Value = 42.0,
            Formula = "=A1+B1",
            Format = "0.00",
            Address = "$C$3",
            Row = 3,
            Col = 3,
            Interior = new Interior(6, 255),
            Font = new CellFont(true, false, 11.0, "Calibri", 0)
        };

        Assert.Equal(6, cell.Color);
        Assert.Equal(255, cell.Rgb);
        Assert.True(cell.Bold);
        Assert.False(cell.Italic);
        Assert.Equal(11.0, cell.FontSize);
    }

    [Fact]
    public void Cell_DefaultValues_AreReasonable()
    {
        var cell = new Cell();

        Assert.Null(cell.Value);
        Assert.Equal("", cell.Formula);
        Assert.Equal("", cell.Format);
        Assert.Equal("", cell.Address);
        Assert.Equal(0, cell.Row);
        Assert.Equal(0, cell.Col);
        Assert.Equal(0, cell.Color);
        Assert.Equal(0, cell.Rgb);
        Assert.False(cell.Bold);
        Assert.False(cell.Italic);
        Assert.Equal(0.0, cell.FontSize);
    }

    [Fact]
    public void Interior_StoresProperties()
    {
        var interior = new Interior(3, 16711680);

        Assert.Equal(3, interior.ColorIndex);
        Assert.Equal(16711680, interior.Color);
    }

    [Fact]
    public void CellFont_StoresProperties()
    {
        var font = new CellFont(true, true, 14.0, "Arial", 255);

        Assert.True(font.Bold);
        Assert.True(font.Italic);
        Assert.Equal(14.0, font.Size);
        Assert.Equal("Arial", font.Name);
        Assert.Equal(255, font.Color);
    }
}
