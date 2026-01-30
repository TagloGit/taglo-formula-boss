using System.Globalization;
using System.Security.Cryptography;
using System.Text;

using FormulaBoss.Parsing;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Transpiles DSL AST to C# source code for UDF generation.
/// </summary>
public class CSharpTranspiler
{
    private readonly HashSet<string> _lambdaParameters = [];
    private bool _requiresObjectModel;

    /// <summary>
    ///     Transpiles a DSL expression to a complete C# UDF class.
    /// </summary>
    /// <param name="expression">The parsed AST expression.</param>
    /// <param name="originalSource">The original DSL source text.</param>
    /// <param name="preferredName">Optional preferred name for the UDF (e.g., from a LET variable).</param>
    /// <returns>The transpilation result with generated code and metadata.</returns>
    public TranspileResult Transpile(Expression expression, string originalSource, string? preferredName = null)
    {
        _requiresObjectModel = false;
        _lambdaParameters.Clear();

        // First pass: detect if object model is needed
        DetectObjectModelUsage(expression);

        // Generate the method name - use preferred name if provided, otherwise hash
        var methodName = GenerateMethodName(originalSource, preferredName);

        // Generate the expression body
        var expressionCode = TranspileExpression(expression);

        // Generate the complete UDF class
        var sourceCode = GenerateUdfClass(methodName, expressionCode);

        return new TranspileResult(sourceCode, methodName, _requiresObjectModel, originalSource);
    }

    private void DetectObjectModelUsage(Expression expression)
    {
        switch (expression)
        {
            case MemberAccess member:
                // .cells triggers object model, .values does not
                if (member.Member is "cells" or "color" or "row" or "col" or "rgb" or "bold" or "italic" or "fontSize"
                    or "format" or "formula" or "address")
                {
                    _requiresObjectModel = true;
                }

                DetectObjectModelUsage(member.Target);
                break;

            case MethodCall call:
                DetectObjectModelUsage(call.Target);
                foreach (var arg in call.Arguments)
                {
                    DetectObjectModelUsage(arg);
                }

                break;

            case BinaryExpr binary:
                DetectObjectModelUsage(binary.Left);
                DetectObjectModelUsage(binary.Right);
                break;

            case UnaryExpr unary:
                DetectObjectModelUsage(unary.Operand);
                break;

            case LambdaExpr lambda:
                DetectObjectModelUsage(lambda.Body);
                break;

            case GroupingExpr grouping:
                DetectObjectModelUsage(grouping.Inner);
                break;
        }
    }

    private string TranspileExpression(Expression expression)
    {
        return expression switch
        {
            IdentifierExpr ident => TranspileIdentifier(ident),
            NumberLiteral num => TranspileNumber(num),
            StringLiteral str => TranspileString(str),
            BinaryExpr binary => TranspileBinary(binary),
            UnaryExpr unary => TranspileUnary(unary),
            MemberAccess member => TranspileMemberAccess(member),
            MethodCall call => TranspileMethodCall(call),
            LambdaExpr lambda => TranspileLambda(lambda),
            GroupingExpr grouping => $"({TranspileExpression(grouping.Inner)})",
            _ => throw new InvalidOperationException($"Unknown expression type: {expression.GetType().Name}")
        };
    }

    private string TranspileIdentifier(IdentifierExpr ident)
    {
        // If it's a lambda parameter, return it directly
        if (_lambdaParameters.Contains(ident.Name))
        {
            return ident.Name;
        }

        // Otherwise it's the input range/data reference
        return "__source__";
    }

    private static string TranspileNumber(NumberLiteral num) => num.Value.ToString(CultureInfo.InvariantCulture);

    private static string TranspileString(StringLiteral str)
    {
        // Escape the string for C#
        var escaped = str.Value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
        return $"\"{escaped}\"";
    }

    private string TranspileBinary(BinaryExpr binary)
    {
        var left = TranspileExpression(binary.Left);
        var right = TranspileExpression(binary.Right);

        // For comparison operators in value-only path, lambda parameters are objects
        // and need to be cast to double for numeric comparisons
        var isComparison = binary.Operator is ">" or "<" or ">=" or "<=" or "==" or "!=";
        if (isComparison && !_requiresObjectModel)
        {
            // Cast lambda parameters to double for numeric comparisons
            if (IsLambdaParameter(binary.Left) && IsNumericLiteral(binary.Right))
            {
                left = $"Convert.ToDouble({left})";
            }

            if (IsLambdaParameter(binary.Right) && IsNumericLiteral(binary.Left))
            {
                right = $"Convert.ToDouble({right})";
            }
        }

        return $"({left} {binary.Operator} {right})";
    }

    private static bool IsNumericLiteral(Expression expr) => expr is NumberLiteral;

    private string TranspileUnary(UnaryExpr unary)
    {
        var operand = TranspileExpression(unary.Operand);
        return $"({unary.Operator}{operand})";
    }

    private string TranspileMemberAccess(MemberAccess member)
    {
        var target = TranspileExpression(member.Target);

        // Special handling for cell properties on lambda parameters
        if (IsLambdaParameter(member.Target))
        {
            return member.Member switch
            {
                "value" => _requiresObjectModel ? $"{target}.Value" : target,
                "color" => $"(int)({target}.Interior.ColorIndex ?? 0)",
                "rgb" => $"(int)({target}.Interior.Color ?? 0)",
                "row" => $"{target}.Row",
                "col" => $"{target}.Column",
                "bold" => $"(bool)({target}.Font.Bold ?? false)",
                "italic" => $"(bool)({target}.Font.Italic ?? false)",
                "fontSize" => $"(double)({target}.Font.Size ?? 11)",
                "format" => $"(string)({target}.NumberFormat ?? \"General\")",
                "formula" => $"(string)({target}.Formula ?? \"\")",
                "address" => $"{target}.Address",
                _ => $"{target}.{member.Member}"
            };
        }

        // Special handling for .cells and .values on the source
        if (target == "__source__")
        {
            return member.Member switch
            {
                "cells" => "__cells__",
                "values" => "__values__",
                _ => $"__source__.{member.Member}"
            };
        }

        return $"{target}.{member.Member}";
    }

    private string TranspileMethodCall(MethodCall call)
    {
        var target = TranspileExpression(call.Target);
        var args = call.Arguments.Select(TranspileExpression).ToList();

        // If calling a LINQ method directly on the source (without .values or .cells),
        // implicitly use .values - e.g., data.where(...) becomes __values__.Where(...)
        if (target == "__source__")
        {
            target = "__values__";
        }

        // Map DSL methods to C#/LINQ equivalents
        return call.Method.ToLowerInvariant() switch
        {
            "where" => $"{target}.Where({string.Join(", ", args)})",
            "select" => $"{target}.Select({string.Join(", ", args)})",
            "toarray" => $"{target}.ToArray()",
            "orderby" => $"{target}.OrderBy({string.Join(", ", args)})",
            "orderbydesc" => $"{target}.OrderByDescending({string.Join(", ", args)})",
            "take" => $"{target}.Take({string.Join(", ", args)})",
            "skip" => $"{target}.Skip({string.Join(", ", args)})",
            "distinct" => $"{target}.Distinct()",
            // For numeric aggregations, cast objects to double
            // On object model path, items might be COM cells - extract .Value if so
            // Use Aggregate() with explicit types to avoid dynamic type inference issues
            "sum" => args.Count > 0
                ? $"{target}.Sum({string.Join(", ", args)})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(0.0, (acc, x) => acc + Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Sum()",
            "avg" or "average" => args.Count > 0
                ? $"{target}.Average({string.Join(", ", args)})"
                : _requiresObjectModel
                    ? $"((Func<double>)(() => {{ var items = {target}.ToList(); return items.Count == 0 ? 0.0 : items.Aggregate(0.0, (acc, x) => acc + Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)) / items.Count; }}))()"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Average()",
            "min" => args.Count > 0
                ? $"{target}.Min({string.Join(", ", args)})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(double.MaxValue, (acc, x) => Math.Min(acc, Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Min()",
            "max" => args.Count > 0
                ? $"{target}.Max({string.Join(", ", args)})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(double.MinValue, (acc, x) => Math.Max(acc, Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Max()",
            "count" => $"{target}.Count()",
            "first" => $"{target}.First()",
            "firstordefault" => $"{target}.FirstOrDefault()",
            "last" => $"{target}.Last()",
            "lastordefault" => $"{target}.LastOrDefault()",
            _ => $"{target}.{call.Method}({string.Join(", ", args)})"
        };
    }

    private string TranspileLambda(LambdaExpr lambda)
    {
        _lambdaParameters.Add(lambda.Parameter);
        var body = TranspileExpression(lambda.Body);
        _lambdaParameters.Remove(lambda.Parameter);
        return $"{lambda.Parameter} => {body}";
    }

    private bool IsLambdaParameter(Expression expression) =>
        expression is IdentifierExpr ident && _lambdaParameters.Contains(ident.Name);

    private string GenerateUdfClass(string methodName, string expressionCode)
    {
        var sb = new StringBuilder();

        sb.AppendLine("using System;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using System.Reflection;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine();
        sb.AppendLine("public static class GeneratedUdf");
        sb.AppendLine("{");

        if (_requiresObjectModel)
        {
            GenerateObjectModelMethod(sb, methodName, expressionCode);
        }
        else
        {
            GenerateValueOnlyMethod(sb, methodName, expressionCode);
        }

        sb.AppendLine("}");

        return sb.ToString();
    }

    private static void GenerateObjectModelMethod(StringBuilder sb, string methodName, string expressionCode)
    {
        // Generate the ExcelDNA entry point - uses reflection to avoid assembly identity issues
        sb.Append("    public static object ").Append(methodName).AppendLine("(object rangeRef)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Get Range via reflection to avoid assembly identity issues");
        sb.AppendLine("            // Object model features (.cells, .color, etc.) require a range reference - they cannot");
        sb.AppendLine("            // work with array output from a previous UDF because cell formatting is not preserved.");
        sb.AppendLine("            if (rangeRef?.GetType()?.Name != \"ExcelReference\")");
        sb.AppendLine("                return \"ERROR: Cell properties require a range reference, not \" + (rangeRef?.GetType()?.Name ?? \"null\") + \". Use a cell range like A1:B10.\";");
        sb.AppendLine();
        sb.AppendLine("            var excelDnaAssembly = rangeRef.GetType().Assembly;");
        sb.AppendLine("            var excelDnaUtilType = excelDnaAssembly.GetType(\"ExcelDna.Integration.ExcelDnaUtil\");");
        sb.AppendLine("            var appProperty = excelDnaUtilType?.GetProperty(\"Application\", BindingFlags.Public | BindingFlags.Static);");
        sb.AppendLine("            dynamic app = appProperty?.GetValue(null);");
        sb.AppendLine("            if (app == null) return \"ERROR: Could not get Excel Application\";");
        sb.AppendLine();
        sb.AppendLine("            // Get range address using ExcelReference's own methods instead of XlCall");
        sb.AppendLine("            var sheetIdProp = rangeRef.GetType().GetProperty(\"SheetId\");");
        sb.AppendLine("            var rowFirstProp = rangeRef.GetType().GetProperty(\"RowFirst\");");
        sb.AppendLine("            var rowLastProp = rangeRef.GetType().GetProperty(\"RowLast\");");
        sb.AppendLine("            var colFirstProp = rangeRef.GetType().GetProperty(\"ColumnFirst\");");
        sb.AppendLine("            var colLastProp = rangeRef.GetType().GetProperty(\"ColumnLast\");");
        sb.AppendLine();
        sb.AppendLine("            var rowFirst = (int)rowFirstProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var rowLast = (int)rowLastProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var colFirst = (int)colFirstProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var colLast = (int)colLastProp.GetValue(rangeRef) + 1;");
        sb.AppendLine();
        sb.AppendLine("            // Convert column numbers to letters");
        sb.AppendLine("            Func<int, string> colToLetter = (col) => {");
        sb.AppendLine("                string result = \"\";");
        sb.AppendLine("                while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }");
        sb.AppendLine("                return result;");
        sb.AppendLine("            };");
        sb.AppendLine();
        sb.AppendLine("            string address;");
        sb.AppendLine("            if (rowFirst == rowLast && colFirst == colLast)");
        sb.AppendLine("                address = colToLetter(colFirst) + rowFirst;");
        sb.AppendLine("            else");
        sb.AppendLine("                address = colToLetter(colFirst) + rowFirst + \":\" + colToLetter(colLast) + rowLast;");
        sb.AppendLine();
        sb.AppendLine("            dynamic range = app.Range[address];");
        sb.Append("            return ").Append(methodName).AppendLine("_Core(range);");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Generate the testable core method (takes Range directly, no ExcelDNA dependencies)
        sb.AppendLine("    /// <summary>");
        sb.AppendLine("    /// Core computation logic - can be called directly with an Excel Range for testing.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static object ").Append(methodName).AppendLine("_Core(dynamic range)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Build cell list manually - COM collections don't work well with LINQ Cast<T>");
        sb.AppendLine("            var __cells__ = new System.Collections.Generic.List<dynamic>();");
        sb.AppendLine("            foreach (dynamic cell in range.Cells) __cells__.Add(cell);");
        sb.AppendLine();

        // Replace placeholders and generate the result
        var code = expressionCode
            .Replace("__source__", "range")
            .Replace("__cells__", "__cells__")
            .Replace("__values__", "range.Value");

        sb.Append("            object result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            // Normalize result inline");
        sb.AppendLine("            if (result == null) return string.Empty;");
        sb.AppendLine("            if (result is string || result is double || result is int || result is bool) return result;");
        sb.AppendLine("            if (result is object[,]) return result;");
        sb.AppendLine("            if (result is System.Collections.IEnumerable enumerable && !(result is string))");
        sb.AppendLine("            {");
        sb.AppendLine("                var list = enumerable.Cast<object>().ToList();");
        sb.AppendLine("                if (list.Count == 0) return string.Empty;");
        sb.AppendLine("                var output = new object[list.Count, 1];");
        sb.AppendLine("                for (var i = 0; i < list.Count; i++)");
        sb.AppendLine("                {");
        sb.AppendLine("                    var item = list[i];");
        sb.AppendLine("                    // If item is a COM cell object, extract its Value");
        sb.AppendLine("                    if (item != null && item.GetType().IsCOMObject)");
        sb.AppendLine("                        output[i, 0] = ((dynamic)item).Value ?? string.Empty;");
        sb.AppendLine("                    else");
        sb.AppendLine("                        output[i, 0] = item ?? string.Empty;");
        sb.AppendLine("                }");
        sb.AppendLine("                return output;");
        sb.AppendLine("            }");
        sb.AppendLine("            return result;");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
    }

    private static void GenerateValueOnlyMethod(StringBuilder sb, string methodName, string expressionCode)
    {
        // Generate the ExcelDNA entry point (uses RuntimeHelpers for ExcelReference → values conversion)
        sb.Append("    public static object ").Append(methodName).AppendLine("(object rangeRef)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Handle both ExcelReference (from range) and object[,] (from previous UDF)");
        sb.AppendLine("            object[,] values;");
        sb.AppendLine("            if (rangeRef?.GetType()?.Name == \"ExcelReference\")");
        sb.AppendLine("            {");
        sb.AppendLine("                // Extract values from ExcelReference via reflection");
        sb.AppendLine("                var getValueMethod = rangeRef.GetType().GetMethod(\"GetValue\", Type.EmptyTypes);");
        sb.AppendLine("                var rawResult = getValueMethod?.Invoke(rangeRef, null);");
        sb.AppendLine("                values = rawResult is object[,] arr ? arr : new object[,] { { rawResult } };");
        sb.AppendLine("            }");
        sb.AppendLine("            else if (rangeRef is object[,] inputArray)");
        sb.AppendLine("            {");
        sb.AppendLine("                // Already an array (from previous UDF in LET chain)");
        sb.AppendLine("                values = inputArray;");
        sb.AppendLine("            }");
        sb.AppendLine("            else if (rangeRef != null)");
        sb.AppendLine("            {");
        sb.AppendLine("                // Single value - wrap in 2D array");
        sb.AppendLine("                values = new object[,] { { rangeRef } };");
        sb.AppendLine("            }");
        sb.AppendLine("            else");
        sb.AppendLine("            {");
        sb.AppendLine("                return \"ERROR: Input is null\";");
        sb.AppendLine("            }");
        sb.AppendLine();
        sb.Append("            return ").Append(methodName).AppendLine("_Core(values);");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Generate the testable core method (takes values directly, no ExcelDNA dependencies)
        sb.AppendLine("    /// <summary>");
        sb.AppendLine("    /// Core computation logic - can be called directly with a values array for testing.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static object ").Append(methodName).AppendLine("_Core(object[,] values)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            var __values__ = values.Cast<object>();");
        sb.AppendLine();

        // Replace placeholders
        var code = expressionCode
            .Replace("__source__", "values")
            .Replace("__values__", "__values__");

        sb.Append("            object result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            // Normalize result inline");
        sb.AppendLine("            if (result == null) return string.Empty;");
        sb.AppendLine("            if (result is string || result is double || result is int || result is bool) return result;");
        sb.AppendLine("            if (result is object[,]) return result;");
        sb.AppendLine("            if (result is System.Collections.IEnumerable enumerable && !(result is string))");
        sb.AppendLine("            {");
        sb.AppendLine("                var list = enumerable.Cast<object>().ToList();");
        sb.AppendLine("                if (list.Count == 0) return string.Empty;");
        sb.AppendLine("                // Always return as 2D array for Excel spill (removed single-item special case)");
        sb.AppendLine("                var output = new object[list.Count, 1];");
        sb.AppendLine("                for (var i = 0; i < list.Count; i++) output[i, 0] = list[i] ?? string.Empty;");
        sb.AppendLine("                return output;");
        sb.AppendLine("            }");
        sb.AppendLine("            return result;");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
    }

    private static string GenerateMethodName(string source, string? preferredName)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            return SanitizeUdfName(preferredName);
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(source));
        var hashString = Convert.ToHexString(hash)[..8].ToLowerInvariant();
        return $"__udf_{hashString}";
    }

    /// <summary>
    ///     Sanitizes a name to be a valid Excel UDF name.
    ///     Converts to uppercase, removes invalid characters, ensures it starts with a letter or underscore.
    /// </summary>
    private static string SanitizeUdfName(string name)
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

        // Must start with letter or underscore
        if (result.Length == 0)
        {
            return "_UDF";
        }

        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        return result;
    }
}
