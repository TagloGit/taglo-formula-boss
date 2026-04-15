using System.Security.Cryptography;
using System.Text;

using FormulaBoss.Compilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Generates C# UDF source code from a <see cref="DetectionResult" />.
///     Generated code references FormulaBoss.Runtime types directly (loaded via host ALC).
/// </summary>
public class CodeEmitter
{
    public const string UdfPrefix = "__FB_";
    public const string DebugSuffix = "_DEBUG";

    private static readonly HashSet<string> ReservedExcelNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "RESULT",
        "SUM",
        "AVERAGE",
        "COUNT",
        "MIN",
        "MAX",
        "IF",
        "AND",
        "OR",
        "NOT",
        "FILTER",
        "SORT",
        "UNIQUE",
        "INDEX",
        "MATCH",
        "VLOOKUP",
        "HLOOKUP",
        "XLOOKUP",
        "LET",
        "LAMBDA",
        "MAP",
        "REDUCE",
        "SCAN",
        "MAKEARRAY",
        "BYROW",
        "BYCOL",
        "CHOOSE",
        "OFFSET",
        "INDIRECT",
        "ROW",
        "COLUMN",
        "ROWS",
        "COLUMNS",
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
        "CONCATENATE",
        "TEXTJOIN",
        "LEFT",
        "RIGHT",
        "MID",
        "LEN",
        "FIND",
        "SEARCH",
        "SUBSTITUTE",
        "REPLACE",
        "TRIM",
        "UPPER",
        "LOWER",
        "PROPER",
        "TEXT",
        "VALUE",
        "TODAY",
        "NOW",
        "DATE",
        "YEAR",
        "MONTH",
        "DAY",
        "HOUR",
        "MINUTE",
        "SECOND",
        "ROUND",
        "ROUNDUP",
        "ROUNDDOWN",
        "INT",
        "MOD",
        "ABS",
        "POWER",
        "SQRT",
        "TRUE",
        "FALSE",
        "NA",
        "ISNA",
        "ISERROR",
        "ISBLANK",
        "ISNUMBER",
        "ISTEXT"
    };

    /// <summary>
    ///     Generates a UDF source code string from a detection result and the original expression.
    /// </summary>
    /// <param name="detection">Detection result from <see cref="InputDetector" />.</param>
    /// <param name="originalExpression">The original user expression.</param>
    /// <param name="preferredName">Optional preferred UDF name.</param>
    /// <param name="headersByParameter">
    ///     Optional column headers per parameter name. When provided, dot-notation column access
    ///     (e.g. <c>r.Population2025</c>) is rewritten to bracket access (<c>r["Population 2025"]</c>).
    /// </param>
    public TranspileResult Emit(
        DetectionResult detection,
        string originalExpression,
        string? preferredName = null,
        Dictionary<string, string[]>? headersByParameter = null)
    {
        var methodName = GenerateMethodName(originalExpression, preferredName);

        // Apply dot-notation-to-bracket rewrite if headers are available
        var rewrittenDetection = detection;
        if (headersByParameter is { Count: > 0 })
        {
            rewrittenDetection = ApplyDotNotationRewrite(detection, headersByParameter);
        }

        var source = BuildSource(rewrittenDetection, methodName);
        return new TranspileResult(source, methodName, detection.RequiresObjectModel, originalExpression);
    }

    /// <summary>
    ///     Emits a debug-instrumented variant of the UDF. The method name gets a
    ///     <see cref="DebugSuffix" /> suffix and the user block is rewritten by
    ///     <see cref="DebugInstrumentationRewriter" /> to fire <c>Tracer</c> calls.
    /// </summary>
    /// <param name="detection">Detection result from <see cref="InputDetector" />.</param>
    /// <param name="originalExpression">The original user expression.</param>
    /// <param name="preferredName">Optional preferred UDF name (suffix is appended).</param>
    /// <param name="headersByParameter">Column header mapping per parameter (as for <see cref="Emit" />).</param>
    /// <param name="callerAddressExpression">
    ///     C# expression that resolves to the caller cell address string. Defaults to an empty
    ///     literal; the compile/register path supplies the real expression.
    /// </param>
    public TranspileResult EmitDebug(
        DetectionResult detection,
        string originalExpression,
        string? preferredName = null,
        Dictionary<string, string[]>? headersByParameter = null,
        string callerAddressExpression = "\"\"")
    {
        var baseName = GenerateMethodName(originalExpression, preferredName);
        var methodName = baseName + DebugSuffix;

        var rewrittenDetection = detection;
        if (headersByParameter is { Count: > 0 })
        {
            rewrittenDetection = ApplyDotNotationRewrite(detection, headersByParameter);
        }

        var instrumentedDetection = InstrumentForDebug(rewrittenDetection, methodName, callerAddressExpression);
        var source = BuildSource(instrumentedDetection, methodName);
        return new TranspileResult(source, methodName, detection.RequiresObjectModel, originalExpression);
    }

    private static DetectionResult InstrumentForDebug(
        DetectionResult detection,
        string traceName,
        string callerAddressExpression)
    {
        // Normalise to a statement block so the rewriter sees a BlockSyntax in every case.
        var blockText = detection.IsStatementBlock
            ? detection.NormalizedExpression
            : $"{{ return {detection.NormalizedExpression}; }}";

        if (SyntaxFactory.ParseStatement(blockText) is not BlockSyntax block)
        {
            return detection;
        }

        var instrumented = DebugInstrumentationRewriter.Instrument(block, traceName, callerAddressExpression);
        return detection with
        {
            NormalizedExpression = instrumented.NormalizeWhitespace().ToFullString(),
            IsStatementBlock = true
        };
    }

    private static DetectionResult ApplyDotNotationRewrite(
        DetectionResult detection,
        Dictionary<string, string[]> headersByParameter)
    {
        // Build a combined column mapping from all parameters that have headers
        var combinedMapping = new Dictionary<string, string>();
        foreach (var (_, headers) in headersByParameter)
        {
            var mapping = ColumnMapper.BuildMapping(headers);
            foreach (var (sanitised, original) in mapping)
            {
                // If same sanitised name maps to different originals across params, skip
                if (combinedMapping.TryGetValue(sanitised, out var existing) && existing != original)
                {
                    combinedMapping.Remove(sanitised);
                }
                else
                {
                    combinedMapping[sanitised] = original;
                }
            }
        }

        if (combinedMapping.Count == 0)
        {
            return detection;
        }

        var headerVarNames = new HashSet<string>(headersByParameter.Keys);
        var rewritten = DotNotationRewriter.Rewrite(
            detection.NormalizedExpression, combinedMapping, headerVarNames);

        return detection with { NormalizedExpression = rewritten };
    }

    internal static string GenerateMethodName(string expression, string? preferredName)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            var sanitized = SanitizeName(preferredName);
            return $"{UdfPrefix}{sanitized}";
        }

        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(expression));
        return $"{UdfPrefix}{Convert.ToHexString(hash)[..8]}";
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

        // Using directives (shared with ImportedTypeNames for free-variable filtering)
        foreach (var ns in ImportedTypeNames.ImportedNamespaces)
        {
            sb.AppendLine($"using {ns};");
        }
        sb.AppendLine();

        // Class
        sb.AppendLine($"public static class {methodName}_Class");
        sb.AppendLine("{");

        // Method signature: all parameters as object
        var allParams = detection.Parameters.Select(p => $"object {p}__raw");
        var paramList = string.Join(", ", allParams);
        sb.AppendLine($"    public static object {methodName}({paramList})");
        sb.AppendLine("    {");

        // Uniform preamble: wrap each parameter
        foreach (var param in detection.Parameters)
        {
            var extractHeaders = detection.HeaderVariables.Contains(param);
            EmitParameterWrapping(sb, param, detection.RequiresObjectModel, extractHeaders);
        }

        if (detection.Parameters.Count > 0)
        {
            sb.AppendLine();
        }

        // User expression body
        if (detection.IsStatementBlock)
        {
            // Statement block: wrap in local function so user's return statements work
            sb.AppendLine("        object __userBlock()");
            sb.AppendLine($"        {detection.NormalizedExpression}");
            sb.AppendLine("        var __result = __userBlock();");
        }
        else
        {
            sb.AppendLine($"        var __result = {detection.NormalizedExpression};");
        }

        sb.AppendLine("        return FormulaBoss.RuntimeHelpers.ToResultDelegate != null");
        sb.AppendLine("            ? FormulaBoss.RuntimeHelpers.ToResultDelegate(__result)");
        sb.AppendLine("            : __result;");

        sb.AppendLine("    }");
        sb.AppendLine("}");

        return sb.ToString();
    }

    private static void EmitParameterWrapping(StringBuilder sb, string param, bool requiresObjectModel,
        bool extractHeaders)
    {
        // Check if parameter is an ExcelReference and extract values
        sb.AppendLine($"        var {param}__isRef = {param}__raw?.GetType()?.Name == \"ExcelReference\";");
        sb.AppendLine($"        var {param}__values = {param}__isRef == true");
        sb.AppendLine($"            ? FormulaBoss.RuntimeHelpers.GetValuesFromReference({param}__raw)");
        sb.AppendLine($"            : {param}__raw;");

        if (extractHeaders)
        {
            // Get headers from first row (only when expression uses r["Col"] syntax for this param)
            sb.AppendLine(
                $"        var {param}__headers = {param}__values is object[,] {param}__arr && {param}__arr.GetLength(0) > 0");
            sb.AppendLine($"            ? FormulaBoss.RuntimeHelpers.GetHeadersDelegate?.Invoke({param}__arr)");
            sb.AppendLine("            : null;");

            // Strip header row from values when headers are present
            sb.AppendLine($"        if ({param}__headers != null && {param}__values is object[,] {param}__valArr)");
            sb.AppendLine("        {");
            sb.AppendLine($"            var __rows = {param}__valArr.GetLength(0) - 1;");
            sb.AppendLine($"            var __cols = {param}__valArr.GetLength(1);");
            sb.AppendLine("            var __stripped = new object[__rows, __cols];");
            sb.AppendLine("            for (var __r = 0; __r < __rows; __r++)");
            sb.AppendLine("                for (var __c = 0; __c < __cols; __c++)");
            sb.AppendLine($"                    __stripped[__r, __c] = {param}__valArr[__r + 1, __c];");
            sb.AppendLine($"            {param}__values = __stripped;");
            sb.AppendLine("        }");
        }
        else
        {
            sb.AppendLine($"        string[]? {param}__headers = null;");
        }

        // Get origin if object model is needed
        if (requiresObjectModel)
        {
            sb.AppendLine($"        var {param}__origin = {param}__isRef == true");
            sb.AppendLine(
                $"            ? (RangeOrigin?)FormulaBoss.RuntimeHelpers.GetOriginDelegate?.Invoke({param}__raw)");
            sb.AppendLine("            : null;");

            // When headers were stripped, the origin must shift down by one row
            // so cell positions align with the data rows, not the header row
            if (extractHeaders)
            {
                sb.AppendLine($"        if ({param}__headers != null && {param}__origin != null)");
                sb.AppendLine(
                    $"            {param}__origin = {param}__origin with {{ TopRow = {param}__origin.TopRow + 1 }};");
            }

            sb.AppendLine(extractHeaders
                ? $"        var {param} = (ExcelTable)ExcelValue.Wrap({param}__values, {param}__headers, {param}__origin);"
                : $"        var {param} = ExcelValue.Wrap({param}__values, {param}__headers, {param}__origin);");
        }
        else
        {
            sb.AppendLine(extractHeaders
                ? $"        var {param} = (ExcelTable)ExcelValue.Wrap({param}__values, {param}__headers);"
                : $"        var {param} = ExcelValue.Wrap({param}__values, {param}__headers);");
        }
    }
}
