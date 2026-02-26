using System.Security.Cryptography;
using System.Text;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Generates C# UDF source code from a <see cref="DetectionResult"/>.
///     Generated code references FormulaBoss.Runtime types directly (loaded via host ALC).
/// </summary>
public class CodeEmitter
{
    private static readonly HashSet<string> ReservedExcelNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "RESULT", "SUM", "AVERAGE", "COUNT", "MIN", "MAX", "IF", "AND", "OR", "NOT",
        "FILTER", "SORT", "UNIQUE", "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "XLOOKUP",
        "LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL",
        "CHOOSE", "OFFSET", "INDIRECT", "ROW", "COLUMN", "ROWS", "COLUMNS",
        "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
        "CONCATENATE", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH",
        "SUBSTITUTE", "REPLACE", "TRIM", "UPPER", "LOWER", "PROPER", "TEXT", "VALUE",
        "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
        "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "MOD", "ABS", "POWER", "SQRT",
        "TRUE", "FALSE", "NA", "ISNA", "ISERROR", "ISBLANK", "ISNUMBER", "ISTEXT"
    };

    /// <summary>
    ///     Generates a UDF source code string from a detection result and the original expression.
    /// </summary>
    public TranspileResult Emit(
        DetectionResult detection,
        string originalExpression,
        string? preferredName = null)
    {
        var methodName = GenerateMethodName(originalExpression, preferredName);
        var source = BuildSource(detection, methodName);
        return new TranspileResult(source, methodName, detection.RequiresObjectModel, originalExpression);
    }

    internal static string GenerateMethodName(string expression, string? preferredName)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            var sanitized = SanitizeName(preferredName);
            return $"__udf_{sanitized}";
        }

        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(expression));
        return $"__udf_{Convert.ToHexString(hash)[..8]}";
    }

    internal static string SanitizeName(string name)
    {
        var upper = name.ToUpperInvariant();
        var sb = new StringBuilder();

        foreach (var c in upper)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
        }

        var result = sb.ToString();
        if (result.Length == 0)
        {
            return "UDF";
        }

        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        if (ReservedExcelNames.Contains(result))
        {
            result = "_" + result;
        }

        return result;
    }

    private static string BuildSource(DetectionResult detection, string methodName)
    {
        var sb = new StringBuilder();

        // Using directives
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using FormulaBoss.Runtime;");
        sb.AppendLine();

        // Class
        sb.AppendLine($"public static class {methodName}_Class");
        sb.AppendLine("{");

        // Method signature: all inputs + free vars as object parameters
        var allParams = new List<string>();
        foreach (var input in detection.Inputs)
        {
            allParams.Add($"object {input}__raw");
        }

        foreach (var freeVar in detection.FreeVariables)
        {
            allParams.Add($"object {freeVar}__raw");
        }

        var paramList = string.Join(", ", allParams);
        sb.AppendLine($"    public static object {methodName}({paramList})");
        sb.AppendLine("    {");

        // Preamble: wrap each input
        // First input is assumed to be a range (cast to IExcelRange)
        // Additional inputs in explicit lambdas may be scalars (keep as ExcelValue)
        for (var i = 0; i < detection.Inputs.Count; i++)
        {
            var isFirstInput = i == 0;
            EmitInputWrapping(sb, detection.Inputs[i], detection.RequiresObjectModel, isFirstInput);
        }

        // Wrap free variables as scalars (keep as ExcelValue, not IExcelRange)
        foreach (var freeVar in detection.FreeVariables)
        {
            sb.AppendLine($"        var {freeVar} = ExcelValue.Wrap({freeVar}__raw);");
        }

        if (detection.Inputs.Count > 0 || detection.FreeVariables.Count > 0)
        {
            sb.AppendLine();
        }

        // User expression body
        EmitExpressionBody(sb, detection);

        sb.AppendLine("    }");
        sb.AppendLine("}");

        return sb.ToString();
    }

    private static void EmitInputWrapping(StringBuilder sb, string input, bool requiresObjectModel,
        bool castToRange)
    {
        // Check if input is an ExcelReference and extract values
        sb.AppendLine($"        var {input}__isRef = {input}__raw?.GetType()?.Name == \"ExcelReference\";");
        sb.AppendLine($"        var {input}__values = {input}__isRef == true");
        sb.AppendLine($"            ? FormulaBoss.RuntimeHelpers.GetValuesFromReference({input}__raw)");
        sb.AppendLine($"            : {input}__raw;");

        // Get headers if available (from ExcelReference or directly from array)
        sb.AppendLine($"        var {input}__headers = {input}__isRef == true");
        sb.AppendLine($"            ? FormulaBoss.RuntimeHelpers.GetHeadersDelegate?.Invoke({input}__raw)");
        sb.AppendLine($"            : ({input}__raw is object[,] {input}__arr && {input}__arr.GetLength(0) > 0");
        sb.AppendLine($"                ? FormulaBoss.RuntimeHelpers.GetHeadersDelegate?.Invoke({input}__raw)");
        sb.AppendLine($"                : null);");

        // Get origin if object model is needed
        if (requiresObjectModel)
        {
            sb.AppendLine($"        var {input}__origin = {input}__isRef == true");
            sb.AppendLine($"            ? (RangeOrigin?)FormulaBoss.RuntimeHelpers.GetOriginDelegate?.Invoke({input}__raw)");
            sb.AppendLine($"            : null;");
            var cast = castToRange ? "(IExcelRange)" : "";
            sb.AppendLine($"        var {input} = {cast}ExcelValue.Wrap({input}__values, {input}__headers, {input}__origin);");
        }
        else
        {
            var cast = castToRange ? "(IExcelRange)" : "";
            sb.AppendLine($"        var {input} = {cast}ExcelValue.Wrap({input}__values, {input}__headers);");
        }
    }

    private static void EmitExpressionBody(StringBuilder sb, DetectionResult detection)
    {
        var expr = detection.NormalizedExpression;

        if (detection.IsSugarSyntax)
        {
            // Sugar: the expression IS the full chain starting with the input identifier
            // e.g., "tbl.Rows.Where(r => r[0] > 5)" — tbl is already a local variable
            sb.AppendLine($"        var __result = {expr};");
        }
        else
        {
            // Explicit lambda: extract the body
            // The expression is "(tbl) => body" or "(tbl, maxVal) => body"
            // We need to emit just the body since parameters are already local variables
            var arrowIndex = FindArrowIndex(expr);
            if (arrowIndex >= 0)
            {
                var body = expr[(arrowIndex + 2)..].TrimStart();
                if (detection.IsStatementBody)
                {
                    // Statement block: emit the block directly but wrap in a helper
                    // Remove outer braces
                    var trimmed = body.Trim();
                    if (trimmed.StartsWith('{') && trimmed.EndsWith('}'))
                    {
                        trimmed = trimmed[1..^1].Trim();
                    }

                    // Emit the block contents directly
                    // Replace 'return x;' with assignment to __result
                    // For simplicity, emit a local function
                    sb.AppendLine("        object __Execute()");
                    sb.AppendLine("        {");
                    sb.AppendLine($"            {trimmed}");
                    sb.AppendLine("        }");
                    sb.AppendLine("        var __result = __Execute();");
                }
                else
                {
                    sb.AppendLine($"        var __result = {body};");
                }
            }
            else
            {
                sb.AppendLine($"        var __result = {expr};");
            }
        }

        sb.AppendLine("        return FormulaBoss.RuntimeHelpers.ToResultDelegate != null");
        sb.AppendLine("            ? FormulaBoss.RuntimeHelpers.ToResultDelegate(__result)");
        sb.AppendLine("            : __result;");
    }

    /// <summary>
    ///     Finds the index of the top-level => arrow in an expression, skipping nested lambdas.
    /// </summary>
    private static int FindArrowIndex(string expr)
    {
        var depth = 0;
        for (var i = 0; i < expr.Length - 1; i++)
        {
            switch (expr[i])
            {
                case '(':
                    depth++;
                    break;
                case ')':
                    depth--;
                    break;
                case '=' when depth == 0 && expr[i + 1] == '>':
                    return i;
            }
        }

        return -1;
    }
}
