using System.Diagnostics;
using System.Reflection;

namespace FormulaBoss;

/// <summary>
/// Runtime helper methods called by generated UDF code.
/// Uses reflection to avoid assembly identity issues between the ExcelDNA
/// assembly loaded by the .xll and the one referenced at compile time.
/// </summary>
public static class RuntimeHelpers
{
    /// <summary>
    /// Converts an ExcelReference to an Excel Range COM object.
    /// Use this for object model access (cell colors, formatting, etc.).
    /// </summary>
    public static dynamic GetRangeFromReference(object rangeRef)
    {
        Debug.WriteLine($"GetRangeFromReference called with: {rangeRef?.GetType()?.FullName ?? "null"}");

        if (rangeRef?.GetType()?.Name != "ExcelReference")
        {
            throw new ArgumentException($"Expected ExcelReference, got {rangeRef?.GetType()?.Name ?? "null"}");
        }

        // Get the ExcelDNA assembly from the rangeRef's type (the actually-loaded assembly)
        var excelDnaAssembly = rangeRef.GetType().Assembly;

        // Get ExcelDnaUtil.Application via reflection
        var excelDnaUtilType = excelDnaAssembly.GetType("ExcelDna.Integration.ExcelDnaUtil");
        var appProperty = excelDnaUtilType?.GetProperty("Application", BindingFlags.Public | BindingFlags.Static);
        dynamic app = appProperty?.GetValue(null)
                      ?? throw new InvalidOperationException("Could not get Excel Application");

        // Get the address using XlCall.Excel(xlfReftext, ref, true) via reflection
        var xlCallType = excelDnaAssembly.GetType("ExcelDna.Integration.XlCall");
        var excelMethod = xlCallType?.GetMethod("Excel", [typeof(int), typeof(object[])]);
        const int xlfReftext = 336; // XlCall.xlfReftext constant

        var address = excelMethod?.Invoke(null, [xlfReftext, new object[] { rangeRef, true }]) as string
                      ?? throw new InvalidOperationException("Could not get range address");

        Debug.WriteLine($"Range address: {address}");
        return app.Range[address];
    }

    /// <summary>
    /// Gets the values from an ExcelReference as a 2D array.
    /// Use this for value-only operations (no object model access needed).
    /// </summary>
    public static object[,] GetValuesFromReference(object rangeRef)
    {
        Debug.WriteLine($"GetValuesFromReference called with: {rangeRef?.GetType()?.FullName ?? "null"}");

        if (rangeRef?.GetType()?.Name != "ExcelReference")
        {
            throw new ArgumentException($"Expected ExcelReference, got {rangeRef?.GetType()?.Name ?? "null"}");
        }

        // Call GetValue() via reflection
        var getValueMethod = rangeRef.GetType().GetMethod("GetValue", Type.EmptyTypes);
        var result = getValueMethod?.Invoke(rangeRef, null);

        Debug.WriteLine($"GetValue returned: {result?.GetType()?.Name ?? "null"}");

        if (result is object[,] arr)
        {
            return arr;
        }

        // Single cell - wrap in 2D array
        return new object[,] { { result } };
    }

    /// <summary>
    /// Normalizes a result for return to Excel.
    /// Handles arrays, enumerables, nulls, and scalars.
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

        if (result is System.Collections.IEnumerable enumerable and not string)
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
