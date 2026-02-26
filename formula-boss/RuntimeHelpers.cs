using System.Collections;
using System.Diagnostics;

namespace FormulaBoss;

/// <summary>
///     Runtime helper methods called by generated UDF code.
///     Uses reflection for ExcelDNA API access to avoid TypeLoadException when
///     this type is loaded from Roslyn-compiled code's assembly context.
/// </summary>
public static class RuntimeHelpers
{
    /// <summary>
    ///     Delegate that resolves an ExcelReference to a COM Range object.
    ///     Initialized by <c>AddIn.AutoOpen</c> during add-in startup with a lambda that calls
    ///     XlCall.Excel(xlfReftext) directly. This uses the "delegate bridge" pattern to work around
    ///     assembly identity constraints:
    ///     <list type="bullet">
    ///         <item>The lambda is JIT-compiled in the host context where ExcelDNA types resolve correctly</item>
    ///         <item>
    ///             This field's type (Func) has no ExcelDNA dependency, so RuntimeHelpers can be loaded
    ///             from Roslyn-compiled code without TypeLoadException
    ///         </item>
    ///         <item>xlfReftext returns a sheet-qualified address, fixing cross-sheet range resolution</item>
    ///     </list>
    ///     See CLAUDE.md "Delegate Bridge Pattern" for full explanation.
    /// </summary>
    public static Func<object, object>? ResolveRangeDelegate { get; set; }

    /// <summary>
    ///     Converts an ExcelReference to an Excel Range COM object.
    ///     Delegates to <see cref="ResolveRangeDelegate" /> which must be set during add-in startup.
    ///     Requires the calling UDF to be registered with IsMacroType = true.
    /// </summary>
    public static object GetRangeFromReference(object? rangeRef)
    {
        if (rangeRef == null)
        {
            return "ERROR: rangeRef is null";
        }

        if (rangeRef.GetType().Name != "ExcelReference")
        {
            return $"ERROR: Expected ExcelReference, got {rangeRef.GetType().Name}";
        }

        if (ResolveRangeDelegate == null)
        {
            return "ERROR: ResolveRangeDelegate not initialized - add-in startup may have failed";
        }

        return ResolveRangeDelegate(rangeRef);
    }

    /// <summary>
    ///     Gets the values from an ExcelReference as a 2D array.
    ///     Use this for value-only operations (no object model access needed).
    /// </summary>
    public static object[,] GetValuesFromReference(object? rangeRef)
    {
        Debug.WriteLine($"GetValuesFromReference called with: {rangeRef?.GetType().FullName ?? "null"}");

        if (rangeRef?.GetType().Name != "ExcelReference")
        {
            throw new ArgumentException($"Expected ExcelReference, got {rangeRef?.GetType().Name ?? "null"}");
        }

        // Call GetValue() via reflection
        var getValueMethod = rangeRef.GetType().GetMethod("GetValue", Type.EmptyTypes);
        var result = getValueMethod?.Invoke(rangeRef, null);

        Debug.WriteLine($"GetValue returned: {result?.GetType().Name ?? "null"}");

        if (result is object[,] arr)
        {
            return arr;
        }

        // Single cell - wrap in 2D array
        return new[,] { { result ?? string.Empty } };
    }

    /// <summary>
    ///     Gets table headers from an ExcelReference via RuntimeBridge.GetHeaders delegate.
    ///     Returns null if the range is not part of an Excel Table or the delegate is not initialized.
    /// </summary>
    public static string[]? GetHeadersFromReference(object? rangeRef)
    {
        if (rangeRef == null || Runtime.RuntimeBridge.GetHeaders == null)
        {
            return null;
        }

        try
        {
            return Runtime.RuntimeBridge.GetHeaders(rangeRef);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"GetHeadersFromReference failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    ///     Gets the range origin (sheet, row, col) from an ExcelReference via RuntimeBridge.GetOrigin delegate.
    ///     Returns null if the delegate is not initialized.
    /// </summary>
    public static Runtime.RangeOrigin? GetOriginFromReference(object? rangeRef)
    {
        if (rangeRef == null || Runtime.RuntimeBridge.GetOrigin == null)
        {
            return null;
        }

        try
        {
            return Runtime.RuntimeBridge.GetOrigin(rangeRef);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"GetOriginFromReference failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    ///     Normalizes a result for return to Excel.
    ///     Handles arrays, enumerables, nulls, and scalars.
    /// </summary>
    public static object NormalizeResult(object? result)
    {
        if (result == null)
        {
            return string.Empty;
        }

        if (result is string)
        {
            return result;
        }

        if (result is object[,])
        {
            return result;
        }

        if (result is IEnumerable enumerable and not string)
        {
            var list = enumerable.Cast<object>().ToList();

            if (list.Count == 0)
            {
                return string.Empty;
            }

            if (list.Count == 1)
            {
                return list[0] ?? string.Empty;
            }

            var output = new object[list.Count, 1];
            for (var i = 0; i < list.Count; i++)
            {
                output[i, 0] = list[i] ?? string.Empty;
            }

            return output;
        }

        return result;
    }
}
