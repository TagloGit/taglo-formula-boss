using ExcelDna.Integration;

using FormulaBoss.Runtime;

namespace FormulaBoss;

/// <summary>
///     Hosts the <c>FB.LastTrace()</c> UDF that spills the most recent debug trace buffer
///     as a 2D array with a header row. Returns an error string when no trace has been captured.
/// </summary>
public static class LastTraceUdf
{
    [ExcelFunction(
        Name = "FB.LastTrace",
        Description = "Returns the most recent debug trace buffer as a spilled table.")]
    public static object LastTrace()
    {
        var buffer = Tracer.LastBuffer;
        if (buffer == null)
        {
            return "#N/A \u2014 no trace captured";
        }

        return buffer.ToObjectArray();
    }
}
