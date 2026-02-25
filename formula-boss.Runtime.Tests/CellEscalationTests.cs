using Xunit;

namespace FormulaBoss.Runtime.Tests;

public class CellEscalationTests : IDisposable
{
    public CellEscalationTests()
    {
        // Set up a mock bridge that returns Cell objects based on position
        RuntimeBridge.GetCell = (sheet, row, col) => new Cell
        {
            Value = $"{sheet}!R{row}C{col}",
            Formula = $"=R{row}C{col}",
            Format = "General",
            Address = $"${(char)('A' + col - 1)}${row}",
            Row = row,
            Col = col,
            Interior = new Interior(row % 8, row * 1000 + col),
            Font = new CellFont(row % 2 == 0, col % 2 == 0, 11.0, "Calibri", 0)
        };
    }

    public void Dispose()
    {
        RuntimeBridge.GetCell = null;
    }

    [Fact]
    public void ColumnValue_Cell_EscalatesToCom()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 100.0, 200.0 }, { 300.0, 400.0 } };
        var array = new ExcelArray(data, origin: origin);

        var firstRow = array.Rows.First();
        var cell = firstRow[0].Cell;

        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Col);
        Assert.Equal("Sheet1!R1C1", cell.Value);
    }

    [Fact]
    public void ColumnValue_Cell_SecondColumn_CorrectPosition()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 100.0, 200.0 } };
        var array = new ExcelArray(data, origin: origin);

        var row = array.Rows.First();
        var cell = row[1].Cell;

        Assert.Equal(1, cell.Row);
        Assert.Equal(2, cell.Col);
    }

    [Fact]
    public void ColumnValue_Cell_SecondRow_CorrectPosition()
    {
        var origin = new RangeOrigin("Sheet1", 5, 3);
        var data = new object?[,] { { 10.0 }, { 20.0 } };
        var array = new ExcelArray(data, origin: origin);

        var secondRow = array.Rows.Skip(1).First();
        var cell = secondRow[0].Cell;

        Assert.Equal(6, cell.Row);
        Assert.Equal(3, cell.Col);
    }

    [Fact]
    public void ColumnValue_Cell_WithoutOrigin_Throws()
    {
        var data = new object?[,] { { 42.0 } };
        var array = new ExcelArray(data);

        var row = array.Rows.First();
        Assert.Throws<InvalidOperationException>(() => row[0].Cell);
    }

    [Fact]
    public void ColumnValue_Cell_WithoutBridge_Throws()
    {
        RuntimeBridge.GetCell = null;
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 42.0 } };
        var array = new ExcelArray(data, origin: origin);

        var row = array.Rows.First();
        Assert.Throws<InvalidOperationException>(() => row[0].Cell);
    }

    [Fact]
    public void ExcelTable_ColumnValue_Cell_WithDynamicAccess()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { "Alice", 100.0 }, { "Bob", 200.0 } };
        var headers = new[] { "Name", "Score" };
        var table = new ExcelTable(data, headers, origin);

        dynamic row = table.Rows.First();
        ColumnValue score = row.Score;
        var cell = score.Cell;

        Assert.Equal(1, cell.Row);
        Assert.Equal(2, cell.Col);
        Assert.False(cell.Bold); // row 1 is odd → not bold
    }

    [Fact]
    public void ExcelArray_Cells_IteratesAllCells()
    {
        var origin = new RangeOrigin("Sheet1", 2, 3);
        var data = new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } };
        var array = new ExcelArray(data, origin: origin);

        var cells = array.Cells.ToList();

        Assert.Equal(4, cells.Count);
        Assert.Equal(2, cells[0].Row);
        Assert.Equal(3, cells[0].Col);
        Assert.Equal(2, cells[1].Row);
        Assert.Equal(4, cells[1].Col);
        Assert.Equal(3, cells[2].Row);
        Assert.Equal(3, cells[2].Col);
        Assert.Equal(3, cells[3].Row);
        Assert.Equal(4, cells[3].Col);
    }

    [Fact]
    public void ExcelArray_Cells_WithoutOrigin_Throws()
    {
        var data = new object?[,] { { 1.0 } };
        var array = new ExcelArray(data);

        Assert.Throws<InvalidOperationException>(() => array.Cells.ToList());
    }

    [Fact]
    public void ExcelScalar_Cells_ReturnsSingleCell()
    {
        var origin = new RangeOrigin("Sheet1", 5, 2);
        var scalar = new ExcelScalar(42.0, origin);

        var cells = scalar.Cells.ToList();

        Assert.Single(cells);
        Assert.Equal(5, cells[0].Row);
        Assert.Equal(2, cells[0].Col);
    }

    [Fact]
    public void ExcelScalar_Cells_WithoutOrigin_Throws()
    {
        var scalar = new ExcelScalar(42.0);

        Assert.Throws<InvalidOperationException>(() => scalar.Cells.ToList());
    }

    [Fact]
    public void ExcelScalar_RowCellEscalation_Works()
    {
        var origin = new RangeOrigin("Sheet1", 3, 1);
        var scalar = new ExcelScalar(42.0, origin);

        var row = scalar.Rows.First();
        var cell = row[0].Cell;

        Assert.Equal(3, cell.Row);
        Assert.Equal(1, cell.Col);
    }

    [Fact]
    public void CellEscalation_InsideWhereLambda_Works()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 10.0 }, { 20.0 }, { 30.0 } };
        var array = new ExcelArray(data, origin: origin);

        // Filter rows where cell color index is 0 (row % 8 == 0 → rows 8, 16, etc. — none here)
        // Row 1: colorIndex = 1%8=1, Row 2: 2%8=2, Row 3: 3%8=3
        var filtered = array.Where(r => r[0].Cell.Color == 2);
        var result = filtered.Rows.ToList();

        Assert.Single(result);
        Assert.Equal(20.0, result[0][0].Value);
    }

    [Fact]
    public void Wrap_WithOrigin_PropagatesToArray()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 42.0 } };

        var wrapped = ExcelValue.Wrap(data, origin: origin);
        var array = Assert.IsType<ExcelArray>(wrapped);
        var cell = array.Rows.First()[0].Cell;

        Assert.Equal(1, cell.Row);
    }

    [Fact]
    public void Wrap_WithOrigin_PropagatesToTable()
    {
        var origin = new RangeOrigin("Sheet1", 1, 1);
        var data = new object?[,] { { 42.0 } };

        var wrapped = ExcelValue.Wrap(data, new[] { "Val" }, origin);
        var table = Assert.IsType<ExcelTable>(wrapped);
        var cell = table.Rows.First()[0].Cell;

        Assert.Equal(1, cell.Row);
    }

    [Fact]
    public void Wrap_WithOrigin_PropagatesToScalar()
    {
        var origin = new RangeOrigin("Sheet1", 3, 5);

        var wrapped = ExcelValue.Wrap(42.0, origin: origin);
        var scalar = Assert.IsType<ExcelScalar>(wrapped);
        var cell = scalar.Cells.First();

        Assert.Equal(3, cell.Row);
        Assert.Equal(5, cell.Col);
    }
}
