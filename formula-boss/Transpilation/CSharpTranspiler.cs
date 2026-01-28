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
    /// <returns>The transpilation result with generated code and metadata.</returns>
    public TranspileResult Transpile(Expression expression, string originalSource)
    {
        _requiresObjectModel = false;
        _lambdaParameters.Clear();

        // First pass: detect if object model is needed
        DetectObjectModelUsage(expression);

        // Generate the method name from a hash of the source
        var methodName = GenerateMethodName(originalSource);

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
        return $"({left} {binary.Operator} {right})";
    }

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
            "sum" => args.Count > 0
                ? $"{target}.Sum({string.Join(", ", args)})"
                : $"{target}.Select(x => Convert.ToDouble(x)).Sum()",
            "avg" or "average" => args.Count > 0
                ? $"{target}.Average({string.Join(", ", args)})"
                : $"{target}.Select(x => Convert.ToDouble(x)).Average()",
            "min" => args.Count > 0
                ? $"{target}.Min({string.Join(", ", args)})"
                : $"{target}.Select(x => Convert.ToDouble(x)).Min()",
            "max" => args.Count > 0
                ? $"{target}.Max({string.Join(", ", args)})"
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
        sb.AppendLine("// Note: Avoiding direct ExcelDna.Integration usage due to assembly identity issues");
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
        // No attributes - registration is handled manually via RegisterDelegates
        sb.Append("    public static object ").Append(methodName).AppendLine("(object rangeRef)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Get Excel Application via reflection to avoid assembly identity issues");
        sb.AppendLine("            var excelDnaUtilType = rangeRef.GetType().Assembly.GetType(\"ExcelDna.Integration.ExcelDnaUtil\");");
        sb.AppendLine("            var appProperty = excelDnaUtilType?.GetProperty(\"Application\");");
        sb.AppendLine("            dynamic app = appProperty?.GetValue(null);");
        sb.AppendLine("            if (app == null) return \"ERROR: Could not get Excel Application\";");
        sb.AppendLine();
        sb.AppendLine("            dynamic range;");
        sb.AppendLine("            var typeName = rangeRef?.GetType()?.Name;");
        sb.AppendLine("            if (typeName == \"ExcelReference\")");
        sb.AppendLine("            {");
        sb.AppendLine("                // Call xlfReftext via reflection to get the address");
        sb.AppendLine("                var xlCallType = rangeRef.GetType().Assembly.GetType(\"ExcelDna.Integration.XlCall\");");
        sb.AppendLine("                var excelMethod = xlCallType?.GetMethod(\"Excel\", new[] { typeof(int), typeof(object[]) });");
        sb.AppendLine("                var xlfReftext = 111; // XlCall.xlfReftext constant");
        sb.AppendLine("                var refText = (string)excelMethod?.Invoke(null, new object[] { xlfReftext, new object[] { rangeRef, true } });");
        sb.AppendLine("                range = app.Range[refText];");
        sb.AppendLine("            }");
        sb.AppendLine("            else");
        sb.AppendLine("            {");
        sb.AppendLine("                return \"ERROR: Expected ExcelReference, got \" + typeName;");
        sb.AppendLine("            }");
        sb.AppendLine();
        sb.AppendLine("            var __cells__ = ((System.Collections.IEnumerable)range.Cells).Cast<dynamic>();");
        sb.AppendLine();

        // Replace placeholders and generate the result
        var code = expressionCode
            .Replace("__source__", "range")
            .Replace("__cells__", "__cells__")
            .Replace("__values__", "range.Value");

        sb.Append("            var result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            return NormalizeResult(result);");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Add the normalize helper
        GenerateNormalizeHelper(sb);
    }

    private static void GenerateValueOnlyMethod(StringBuilder sb, string methodName, string expressionCode)
    {
        // No attributes needed - we register manually via RegisterDelegates
        sb.Append("    public static object ").Append(methodName).AppendLine("(object rangeRef)");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Use reflection to avoid assembly mismatch with ExcelReference");
        sb.AppendLine("            object[,] values;");
        sb.AppendLine("            var typeName = rangeRef?.GetType()?.Name;");
        sb.AppendLine("            if (typeName == \"ExcelReference\")");
        sb.AppendLine("            {");
        sb.AppendLine("                // Call GetValue via reflection to avoid type mismatch");
        sb.AppendLine("                var getValueMethod = rangeRef.GetType().GetMethod(\"GetValue\");");
        sb.AppendLine("                var val = getValueMethod?.Invoke(rangeRef, null);");
        sb.AppendLine("                if (val is object[,] arr)");
        sb.AppendLine("                    values = arr;");
        sb.AppendLine("                else");
        sb.AppendLine("                    values = new object[1, 1] { { val } };");
        sb.AppendLine("            }");
        sb.AppendLine("            else if (rangeRef is object[,] arr)");
        sb.AppendLine("            {");
        sb.AppendLine("                values = arr;");
        sb.AppendLine("            }");
        sb.AppendLine("            else");
        sb.AppendLine("            {");
        sb.AppendLine("                values = new object[1, 1] { { rangeRef } };");
        sb.AppendLine("            }");
        sb.AppendLine();
        sb.AppendLine("            var __values__ = values.Cast<object>();");
        sb.AppendLine();

        // Replace placeholders
        var code = expressionCode
            .Replace("__source__", "values")
            .Replace("__values__", "__values__");

        sb.Append("            var result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            return NormalizeResult(result);");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Add the normalize helper
        GenerateNormalizeHelper(sb);
    }

    private static void GenerateNormalizeHelper(StringBuilder sb)
    {
        sb.AppendLine("    private static object NormalizeResult(object result)");
        sb.AppendLine("    {");
        sb.AppendLine("        if (result == null) return \"\";");
        sb.AppendLine("        if (result is string) return result;");
        sb.AppendLine("        if (result is object[,]) return result;");
        sb.AppendLine("        if (result is System.Collections.IEnumerable enumerable)");
        sb.AppendLine("        {");
        sb.AppendLine("            var list = enumerable.Cast<object>().ToList();");
        sb.AppendLine("            if (list.Count == 0) return \"\";");
        sb.AppendLine("            if (list.Count == 1) return list[0] ?? \"\";");
        sb.AppendLine("            var output = new object[list.Count, 1];");
        sb.AppendLine("            for (int i = 0; i < list.Count; i++)");
        sb.AppendLine("                output[i, 0] = list[i] ?? \"\";");
        sb.AppendLine("            return output;");
        sb.AppendLine("        }");
        sb.AppendLine("        return result;");
        sb.AppendLine("    }");
    }

    private static string GenerateMethodName(string source)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(source));
        var hashString = Convert.ToHexString(hash)[..8].ToLowerInvariant();
        return $"__udf_{hashString}";
    }
}
