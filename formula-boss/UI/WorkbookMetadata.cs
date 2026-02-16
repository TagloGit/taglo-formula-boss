using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FormulaBoss.UI;

/// <summary>
///     Snapshot of workbook metadata (table names, named ranges, column headers)
///     captured on the Excel thread for use by the completion system.
/// </summary>
public record WorkbookMetadata(
    IReadOnlyList<string> TableNames,
    IReadOnlyList<string> NamedRanges,
    IReadOnlyDictionary<string, IReadOnlyList<string>> TableColumns)
{
    public static readonly WorkbookMetadata Empty = new(
        Array.Empty<string>(),
        Array.Empty<string>(),
        new Dictionary<string, IReadOnlyList<string>>());

    /// <summary>
    ///     Captures table names, named ranges, and column headers from the active workbook.
    ///     Must be called on the Excel thread. All COM objects are released.
    /// </summary>
    public static WorkbookMetadata CaptureFromExcel(dynamic app)
    {
        var tableNames = new List<string>();
        var namedRanges = new List<string>();
        var tableColumns = new Dictionary<string, IReadOnlyList<string>>();

        dynamic? workbook = null;
        try
        {
            workbook = app.ActiveWorkbook;
            if (workbook == null)
            {
                return Empty;
            }

            CaptureNamedRanges(workbook, namedRanges);
            CaptureTables(workbook, tableNames, tableColumns);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"WorkbookMetadata.CaptureFromExcel error: {ex.Message}");
        }
        finally
        {
            ReleaseCom(workbook);
        }

        return new WorkbookMetadata(tableNames, namedRanges, tableColumns);
    }

    private static void CaptureNamedRanges(dynamic workbook, List<string> namedRanges)
    {
        dynamic? names = null;
        try
        {
            names = workbook.Names;
            int count = names.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic? name = null;
                try
                {
                    name = names.Item(i);

                    // Skip hidden names (Excel internal definitions)
                    bool visible = name.Visible;
                    if (!visible)
                    {
                        continue;
                    }

                    var nameStr = name.Name as string;
                    if (!string.IsNullOrEmpty(nameStr))
                    {
                        // Strip sheet qualifier for local names (Sheet1!MyRange -> MyRange)
                        var excl = nameStr.IndexOf('!');
                        if (excl >= 0)
                        {
                            nameStr = nameStr[(excl + 1)..];
                        }

                        namedRanges.Add(nameStr);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"CaptureNamedRanges item error: {ex.Message}");
                }
                finally
                {
                    ReleaseCom(name);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"CaptureNamedRanges error: {ex.Message}");
        }
        finally
        {
            ReleaseCom(names);
        }
    }

    private static void CaptureTables(dynamic workbook, List<string> tableNames,
        Dictionary<string, IReadOnlyList<string>> tableColumns)
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            int sheetCount = sheets.Count;
            for (var s = 1; s <= sheetCount; s++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;
                try
                {
                    sheet = sheets[s];
                    listObjects = sheet.ListObjects;
                    int loCount = listObjects.Count;
                    for (var t = 1; t <= loCount; t++)
                    {
                        dynamic? lo = null;
                        dynamic? listColumns = null;
                        try
                        {
                            lo = listObjects[t];
                            var tableName = lo.Name as string;
                            if (string.IsNullOrEmpty(tableName))
                            {
                                continue;
                            }

                            tableNames.Add(tableName);

                            var cols = new List<string>();
                            listColumns = lo.ListColumns;
                            int colCount = listColumns.Count;
                            for (var c = 1; c <= colCount; c++)
                            {
                                dynamic? col = null;
                                try
                                {
                                    col = listColumns[c];
                                    var colName = col.Name as string;
                                    if (!string.IsNullOrEmpty(colName))
                                    {
                                        cols.Add(colName);
                                    }
                                }
                                finally
                                {
                                    ReleaseCom(col);
                                }
                            }

                            tableColumns[tableName] = cols;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"CaptureTables table error: {ex.Message}");
                        }
                        finally
                        {
                            ReleaseCom(listColumns);
                            ReleaseCom(lo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"CaptureTables sheet error: {ex.Message}");
                }
                finally
                {
                    ReleaseCom(listObjects);
                    ReleaseCom(sheet);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"CaptureTables error: {ex.Message}");
        }
        finally
        {
            ReleaseCom(sheets);
        }
    }

    private static void ReleaseCom(object? obj)
    {
        if (obj != null)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                /* already released */
            }
        }
    }
}
