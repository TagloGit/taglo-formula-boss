namespace FormulaBoss.Runtime;

public interface IExcelRange
{
    IEnumerable<Row> Rows { get; }

    /// <summary>
    ///     Iterates all cells in the range as <see cref="Cell" /> objects.
    ///     Forces IsMacroType — requires positional context and <see cref="RuntimeBridge.GetCell" />.
    /// </summary>
    IEnumerable<Cell> Cells { get; }

    // Element-wise operations
    IExcelRange Where(Func<Row, bool> predicate);
    IExcelRange Select(Func<Row, ExcelValue> selector);
    IExcelRange SelectMany(Func<Row, IEnumerable<ExcelValue>> selector);
    bool Any(Func<Row, bool> predicate);
    bool All(Func<Row, bool> predicate);
    ExcelValue First(Func<Row, bool> predicate);
    ExcelValue? FirstOrDefault(Func<Row, bool> predicate);

    // Aggregations
    int Count();
    ExcelScalar Sum();
    ExcelScalar Min();
    ExcelScalar Max();
    ExcelScalar Average();

    // Shape-preserving transform
    IExcelRange Map(Func<Row, ExcelValue> selector);

    // Sorting
    IExcelRange OrderBy(Func<Row, object> keySelector);
    IExcelRange OrderByDescending(Func<Row, object> keySelector);

    // Partitioning
    IExcelRange Take(int count);
    IExcelRange Skip(int count);
    IExcelRange Distinct();

    // Folding
    ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func);
    IExcelRange Scan(ExcelValue seed, Func<ExcelValue, Row, ExcelValue> func);
}
