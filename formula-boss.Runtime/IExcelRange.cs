namespace FormulaBoss.Runtime;

public interface IExcelRange
{
    RowCollection Rows { get; }

    /// <summary>
    ///     Iterates all cells in the range as <see cref="Cell" /> objects.
    ///     Forces IsMacroType — requires positional context and <see cref="RuntimeBridge.GetCell" />.
    ///     Only available on ranges with a <see cref="RangeOrigin" /> — transformation operators
    ///     (Where, Select, OrderBy, etc.) drop origin because row positions no longer map to
    ///     original worksheet cells. Calling this on a transformed range throws
    ///     <see cref="InvalidOperationException" />.
    /// </summary>
    IEnumerable<Cell> Cells { get; }

    // Element-wise operations (lambda receives each cell as ExcelValue)
    IExcelRange Where(Func<ExcelValue, bool> predicate);
    IExcelRange Select(Func<ExcelValue, ExcelValue> selector);
    IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector);
    bool Any(Func<ExcelValue, bool> predicate);
    bool All(Func<ExcelValue, bool> predicate);
    ExcelValue First(Func<ExcelValue, bool> predicate);
    ExcelValue? FirstOrDefault(Func<ExcelValue, bool> predicate);

    // Aggregations
    int Count();
    ExcelScalar Sum();
    ExcelScalar Min();
    ExcelScalar Max();
    ExcelScalar Average();

    // Shape-preserving transform
    IExcelRange Map(Func<ExcelValue, ExcelValue> selector);

    // Sorting
    IExcelRange OrderBy(Func<ExcelValue, object> keySelector);
    IExcelRange OrderByDescending(Func<ExcelValue, object> keySelector);

    // Partitioning
    IExcelRange Take(int count);
    IExcelRange Skip(int count);
    IExcelRange Distinct();

    // Folding
    ExcelValue Aggregate(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);
    IExcelRange Scan(ExcelValue seed, Func<ExcelValue, ExcelValue, ExcelValue> func);
}
