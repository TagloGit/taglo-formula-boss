namespace FormulaBoss.Runtime;

/// <summary>Represents a range of Excel values supporting element-wise operations and aggregations.</summary>
public interface IExcelRange
{
    /// <summary>Gets the rows of this range as a <see cref="RowCollection" />.</summary>
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

    /// <summary>Filters elements to those matching the predicate.</summary>
    /// <param name="predicate">A function that returns true for elements to keep.</param>
    /// <returns>A new range containing only matching elements.</returns>
    IExcelRange Where(Func<ExcelValue, bool> predicate);

    /// <summary>Projects each element into a new value.</summary>
    /// <param name="selector">A function that transforms each element.</param>
    /// <returns>A new range containing the transformed values.</returns>
    IExcelRange Select(Func<ExcelValue, ExcelValue> selector);

    /// <summary>Projects each element to a sequence and flattens the results.</summary>
    /// <param name="selector">A function that returns a sequence for each element.</param>
    /// <returns>A new range containing all flattened values.</returns>
    IExcelRange SelectMany(Func<ExcelValue, IEnumerable<ExcelValue>> selector);

    /// <summary>Returns true if any element matches the predicate.</summary>
    /// <param name="predicate">A function to test each element.</param>
    bool Any(Func<ExcelValue, bool> predicate);

    /// <summary>Returns true if all elements match the predicate.</summary>
    /// <param name="predicate">A function to test each element.</param>
    bool All(Func<ExcelValue, bool> predicate);

    /// <summary>Returns the first element matching the predicate, or throws if none found.</summary>
    /// <param name="predicate">A function to test each element.</param>
    ExcelValue First(Func<ExcelValue, bool> predicate);

    /// <summary>Returns the first element matching the predicate, or null if none found.</summary>
    /// <param name="predicate">A function to test each element.</param>
    ExcelValue? FirstOrDefault(Func<ExcelValue, bool> predicate);

    /// <summary>Returns the number of elements in this range.</summary>
    int Count();

    /// <summary>Returns the sum of all elements (converted to double).</summary>
    ExcelScalar Sum();

    /// <summary>Returns the minimum value across all elements.</summary>
    ExcelScalar Min();

    /// <summary>Returns the maximum value across all elements.</summary>
    ExcelScalar Max();

    /// <summary>Returns the arithmetic mean of all elements.</summary>
    ExcelScalar Average();

    /// <summary>Applies a function to each element, preserving the original 2D shape.</summary>
    /// <param name="selector">A function that transforms each element.</param>
    /// <returns>A new range with the same dimensions, containing transformed values.</returns>
    IExcelRange Map(Func<ExcelValue, ExcelValue> selector);

    /// <summary>Sorts elements in ascending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each element.</param>
    IExcelRange OrderBy(Func<ExcelValue, object> keySelector);

    /// <summary>Sorts elements in descending order by the selected key.</summary>
    /// <param name="keySelector">A function that extracts a sort key from each element.</param>
    IExcelRange OrderByDescending(Func<ExcelValue, object> keySelector);

    /// <summary>Returns the first <paramref name="count" /> elements. If negative, returns the last N elements.</summary>
    /// <param name="count">Number of elements to take.</param>
    IExcelRange Take(int count);

    /// <summary>Skips the first <paramref name="count" /> elements. If negative, skips the last N elements.</summary>
    /// <param name="count">Number of elements to skip.</param>
    IExcelRange Skip(int count);

    /// <summary>Returns distinct elements (compared by string representation).</summary>
    IExcelRange Distinct();

    /// <summary>Executes an action for each element in the range.</summary>
    /// <param name="action">The action to perform on each element.</param>
    void ForEach(Action<ExcelValue> action);

    /// <summary>Applies an accumulator function over the elements, returning the final result.</summary>
    /// <param name="seed">The initial accumulator value.</param>
    /// <param name="func">A function that takes (accumulator, element) and returns the new accumulator.</param>
    dynamic Aggregate(dynamic seed, Func<dynamic, dynamic, dynamic> func);

    /// <summary>Like Aggregate, but returns all intermediate accumulator values as a range.</summary>
    /// <param name="seed">The initial accumulator value.</param>
    /// <param name="func">A function that takes (accumulator, element) and returns the new accumulator.</param>
    IExcelRange Scan(dynamic seed, Func<dynamic, dynamic, dynamic> func);
}
