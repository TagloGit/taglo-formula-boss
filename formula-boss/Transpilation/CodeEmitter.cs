using System.Text;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Generates C# source code for a UDF from an <see cref="InputDetectionResult" />.
///     The generated code wraps inputs with <c>ExcelValue.Wrap()</c> and passes
///     the user's expression through mostly verbatim.
/// </summary>
public static class CodeEmitter
{
    private static readonly HashSet<string> ReservedExcelNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "AND", "OR", "NOT",
        "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "CHOOSE", "INDIRECT", "OFFSET",
        "LEFT", "RIGHT", "MID", "LEN", "TRIM", "CLEAN", "UPPER", "LOWER", "PROPER",
        "SUBSTITUTE", "REPLACE", "FIND", "SEARCH", "TEXT", "VALUE", "FIXED", "DOLLAR",
        "EXACT", "REPT", "CHAR", "CODE", "CONCAT", "TEXTJOIN", "NUMBERVALUE",
        "ABS", "INT", "MOD", "ROUND", "ROUNDUP", "ROUNDDOWN", "RAND", "RANDBETWEEN",
        "NOW", "TODAY", "DATE", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
        "DATEDIF", "EDATE", "EOMONTH", "NETWORKDAYS", "WORKDAY", "WEEKDAY", "WEEKNUM",
        "TRUE", "FALSE", "TYPE", "N", "T", "NA",
        "SORT", "SORTBY", "FILTER", "UNIQUE", "SEQUENCE", "RANDARRAY",
        "XLOOKUP", "XMATCH", "LET", "LAMBDA", "SWITCH", "IFS",
        "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "VALUETOTEXT",
        "ROWS", "COLUMNS", "TRANSPOSE", "AREAS",
        "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
        "LARGE", "SMALL", "RANK", "PERCENTILE", "QUARTILE",
        "COUNTA", "COUNTBLANK", "PRODUCT", "SUMPRODUCT",
        "LOG", "LOG10", "LN", "EXP", "POWER", "SQRT", "SIGN", "CEILING", "FLOOR",
        "PI", "FACT", "COMBIN", "PERMUT", "GCD", "LCM",
        "CONCATENATE", "ADDRESS", "CELL", "INFO", "ISTEXT", "ISNUMBER",
        "ISBLANK", "ISERROR", "ISNA", "ISLOGICAL", "ISREF", "ISEVEN", "ISODD",
        "HYPERLINK", "FORMULATEXT",
        "AGGREGATE", "SUBTOTAL",
        "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL",
        "TAKE", "DROP", "EXPAND", "WRAPCOLS", "WRAPROWS", "TOCOL", "TOROW",
        "HSTACK", "VSTACK", "CHOOSEROWS", "CHOOSECOLS",
    };

    /// <summary>
    ///     Generates a complete C# source file containing a static UDF class.
    /// </summary>
    /// <param name="detection">The input detection result.</param>
    /// <param name="methodName">The sanitized UDF method name.</param>
    /// <param name="originalExpression">The original expression (for comments).</param>
    /// <returns>A <see cref="TranspileResult" /> with the generated source.</returns>
    public static TranspileResult Emit(InputDetectionResult detection, string methodName, string originalExpression)
    {
        var sanitized = SanitizeUdfName(methodName);
        var sb = new StringBuilder();

        sb.AppendLine("using System;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using FormulaBoss.Runtime;");
        sb.AppendLine();
        sb.AppendLine("namespace FormulaBoss.Generated");
        sb.AppendLine("{");
        sb.AppendLine($"    public static class {sanitized}_Class");
        sb.AppendLine("    {");

        EmitUdfMethod(sb, detection, sanitized);

        sb.AppendLine("    }");
        sb.AppendLine("}");

        // Collect free variables that were used as additional parameters
        IReadOnlyList<string>? usedColumnBindings = detection.FreeVariables.Count > 0
            ? detection.FreeVariables
            : null;

        return new TranspileResult(sb.ToString(), sanitized, detection.RequiresObjectModel,
            originalExpression, usedColumnBindings);
    }

    private static void EmitUdfMethod(StringBuilder sb, InputDetectionResult detection, string methodName)
    {
        // Build parameter list: all inputs + free variables are object parameters
        var allParams = new List<string>();

        foreach (var input in detection.Inputs)
        {
            allParams.Add($"object {EscapeIdentifier(input)}__raw");
        }

        foreach (var freeVar in detection.FreeVariables)
        {
            allParams.Add($"object {EscapeIdentifier(freeVar)}__raw");
        }

        var paramList = string.Join(", ", allParams);

        sb.AppendLine($"        public static object {methodName}({paramList})");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"hello\";");
        sb.AppendLine("            try");
        sb.AppendLine("            {");

        // Wrap each input with ExcelValue.Wrap()
        // For inputs, resolve ExcelReference to values first using RuntimeHelpers
        foreach (var input in detection.Inputs)
        {
            var escaped = EscapeIdentifier(input);
            EmitWrapInput(sb, escaped);
        }

        foreach (var freeVar in detection.FreeVariables)
        {
            var escaped = EscapeIdentifier(freeVar);
            EmitWrapInput(sb, escaped);
        }

        sb.AppendLine();

        // Emit the user's code
        if (detection.IsStatementBody)
        {
            EmitStatementBody(sb, detection);
        }
        else
        {
            EmitExpressionBody(sb, detection);
        }

        sb.AppendLine("            }");
        sb.AppendLine("            catch (Exception ex)");
        sb.AppendLine("            {");
        sb.AppendLine("                return new object[,] { { $\"ERROR: {ex.GetType().Name}: {ex.Message}\" } };");
        sb.AppendLine("            }");
        sb.AppendLine("        }");
    }

    private static void EmitWrapInput(StringBuilder sb, string name)
    {
        // Resolve ExcelReference to raw values, headers, and origin using RuntimeHelpers
        sb.AppendLine($"            var {name}__isRef = {name}__raw?.GetType()?.Name == \"ExcelReference\";");
        sb.AppendLine($"            var {name}__values = {name}__isRef == true");
        sb.AppendLine($"                ? FormulaBoss.RuntimeHelpers.GetValuesFromReference({name}__raw)");
        sb.AppendLine($"                : {name}__raw;");
        sb.AppendLine($"            var {name}__headers = {name}__isRef == true");
        sb.AppendLine($"                ? FormulaBoss.RuntimeHelpers.GetHeadersFromReference({name}__raw)");
        sb.AppendLine($"                : null;");
        sb.AppendLine($"            var {name}__origin = {name}__isRef == true");
        sb.AppendLine($"                ? FormulaBoss.RuntimeHelpers.GetOriginFromReference({name}__raw)");
        sb.AppendLine($"                : null;");
        sb.AppendLine($"            var {name} = ExcelValue.Wrap({name}__values, {name}__headers, {name}__origin);");
    }

    private static void EmitExpressionBody(StringBuilder sb, InputDetectionResult detection)
    {
        var body = detection.Body;
        // Cast to object first to avoid dynamic dispatch issues with extension methods
        sb.AppendLine("            return new object[,] { { \"hello\" } };");
    }

    private static void EmitStatementBody(StringBuilder sb, InputDetectionResult detection)
    {
        // For statement bodies, we need to wrap the block in a way that captures the return value.
        // The body is a { ... } block with return statements.
        // We use a local function approach.
        sb.AppendLine("            object __exec()");
        sb.AppendLine($"            {detection.Body}");
        sb.AppendLine("            return FormulaBoss.Runtime.ResultConverter.ToResult(__exec());");
    }

    /// <summary>
    ///     Sanitizes a name for use as a UDF method name.
    ///     Converts to uppercase, removes non-alphanumeric chars, prefixes if needed.
    /// </summary>
    internal static string SanitizeUdfName(string name)
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
            return "_UDF";
        }

        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        if (IsReservedExcelName(result))
        {
            result = "_" + result;
        }

        return result;
    }

    internal static bool IsReservedExcelName(string name) =>
        ReservedExcelNames.Contains(name);

    private static string EscapeIdentifier(string name)
    {
        // C# keywords that might be used as variable names
        return name switch
        {
            "class" or "static" or "void" or "return" or "if" or "else" or "for" or "while"
                or "do" or "switch" or "case" or "break" or "continue" or "new" or "this"
                or "base" or "true" or "false" or "null" or "string" or "int" or "double"
                or "bool" or "object" or "decimal" or "float" or "long" or "short" or "byte"
                or "var" or "using" or "namespace" => $"@{name}",
            _ => name
        };
    }
}
