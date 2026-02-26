using System.Text.RegularExpressions;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Result of analyzing a user expression for inputs, object model usage, and free variables.
/// </summary>
/// <param name="IsSugarSyntax">True if the expression is sugar (implicit single input), not an explicit lambda.</param>
/// <param name="Inputs">Input identifier names (e.g., ["tbl"] or ["tbl", "maxVal"]).</param>
/// <param name="RequiresObjectModel">True if .Cell or .Cells is used (needs IsMacroType).</param>
/// <param name="IsStatementBody">True if the expression body is a { ... } block.</param>
/// <param name="FreeVariables">Identifiers not in Inputs or lambda params — become extra UDF parameters.</param>
/// <param name="HasStringBracketAccess">True if r["Col Name"] syntax is detected.</param>
/// <param name="NormalizedExpression">Expression with range refs replaced by valid C# identifiers.</param>
/// <param name="RangeRefMap">Maps placeholder identifiers back to original range refs (e.g., "__range_A1_B10" → "A1:B10").</param>
public record DetectionResult(
    bool IsSugarSyntax,
    IReadOnlyList<string> Inputs,
    bool RequiresObjectModel,
    bool IsStatementBody,
    IReadOnlyList<string> FreeVariables,
    bool HasStringBracketAccess,
    string NormalizedExpression,
    IReadOnlyDictionary<string, string> RangeRefMap);

/// <summary>
///     Analyzes user expressions using Roslyn syntax trees to detect inputs,
///     object model usage, free variables, and range references.
/// </summary>
public class InputDetector
{
    // Matches Excel range references: A1:B10, $A$1:$B$10, A1, $A$1, etc.
    // Must appear before a dot or at end of string to avoid false positives
    private static readonly Regex RangeRefPattern = new(
        @"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?",
        RegexOptions.Compiled);

    // C# keywords and common type names that should not be treated as free variables
    private static readonly HashSet<string> IgnoredIdentifiers = new(StringComparer.Ordinal)
    {
        // C# keywords commonly used in expressions
        "true", "false", "null", "new", "typeof", "nameof", "default", "is", "as", "in",
        "var", "return", "if", "else", "for", "foreach", "while", "do", "switch", "case",
        "break", "continue", "throw", "try", "catch", "finally",
        // Common type names
        "string", "int", "double", "float", "bool", "object", "decimal", "long", "char", "byte",
        "Convert", "Math", "String", "Int32", "Double", "Boolean",
        "Enumerable", "System", "Linq",
        // Runtime type names (these resolve at runtime, not as free vars)
        "ExcelValue", "ExcelArray", "ExcelTable", "ExcelScalar", "Row", "ColumnValue",
        "Cell", "Interior", "CellFont", "RangeOrigin", "ResultConverter",
        "IExcelRange", "RuntimeBridge"
    };

    /// <summary>
    ///     Analyzes a user expression and returns detection results.
    /// </summary>
    public DetectionResult Detect(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            throw new TranspileException("Expression is empty");
        }

        // Step 1: Pre-process range references
        var (normalized, rangeRefMap) = PreprocessRangeRefs(expression);

        // Step 2: Determine if this is a lambda or sugar syntax
        // Wrap in a method body so Roslyn can parse it as an expression
        var wrappedSource = $"class __Wrapper {{ object __M() => {normalized}; }}";
        var tree = CSharpSyntaxTree.ParseText(wrappedSource);
        var root = tree.GetRoot();

        // Find the expression node
        var methodDecl = root.DescendantNodes().OfType<MethodDeclarationSyntax>().First();
        var exprBody = methodDecl.ExpressionBody?.Expression;

        if (exprBody == null)
        {
            throw new TranspileException($"Failed to parse expression: {expression}");
        }

        // Step 3: Detect sugar vs lambda
        var lambdaParams = new List<string>();
        var isStatementBody = false;
        ExpressionSyntax bodyExpression;

        if (exprBody is ParenthesizedLambdaExpressionSyntax parenLambda)
        {
            // Explicit lambda: (a, b) => ...
            lambdaParams.AddRange(parenLambda.ParameterList.Parameters.Select(p => p.Identifier.Text));
            isStatementBody = parenLambda.Block != null;
            bodyExpression = exprBody;
        }
        else if (exprBody is SimpleLambdaExpressionSyntax simpleLambda)
        {
            // Single-param lambda: x => ...
            lambdaParams.Add(simpleLambda.Parameter.Identifier.Text);
            isStatementBody = simpleLambda.Block != null;
            bodyExpression = exprBody;
        }
        else
        {
            // Sugar syntax: tbl.Rows.Where(...)
            bodyExpression = exprBody;
        }

        var isSugar = lambdaParams.Count == 0;
        List<string> inputs;

        if (isSugar)
        {
            // Extract the primary input: leftmost identifier in the leading member-access chain
            var primaryInput = ExtractPrimaryInput(exprBody);
            if (primaryInput == null)
            {
                throw new TranspileException($"Could not detect input identifier in: {expression}");
            }

            inputs = [primaryInput];
        }
        else
        {
            inputs = lambdaParams;
        }

        // Step 4: Detect object model usage (.Cell, .Cells)
        var requiresObjectModel = DetectObjectModel(root);

        // Step 5: Detect free variables
        var allLambdaParams = CollectAllLambdaParameters(root);
        var freeVars = DetectFreeVariables(root, inputs, allLambdaParams);

        // Step 6: Detect string bracket access
        var hasStringBracket = DetectStringBracketAccess(root);

        return new DetectionResult(
            isSugar,
            inputs,
            requiresObjectModel,
            isStatementBody,
            freeVars,
            hasStringBracket,
            normalized,
            rangeRefMap);
    }

    /// <summary>
    ///     Replaces Excel range references with valid C# identifiers.
    /// </summary>
    internal static (string Normalized, IReadOnlyDictionary<string, string> Map) PreprocessRangeRefs(
        string expression)
    {
        var map = new Dictionary<string, string>();
        var result = RangeRefPattern.Replace(expression, match =>
        {
            var rangeRef = match.Value;

            // Only treat as range ref if it contains a colon (A1:B10) or starts with $
            // Single identifiers like A1 could be valid C# identifiers used as variable names
            if (!rangeRef.Contains(':') && !rangeRef.StartsWith('$'))
            {
                return rangeRef;
            }

            var placeholder = "__range_" + rangeRef
                .Replace(":", "_")
                .Replace("$", "");

            map[placeholder] = rangeRef;
            return placeholder;
        });

        return (result, map);
    }

    /// <summary>
    ///     Extracts the primary input identifier from the leftmost position in a member-access chain.
    /// </summary>
    private static string? ExtractPrimaryInput(ExpressionSyntax expr)
    {
        // Walk down the left side of member access / invocation chains
        var current = expr;
        while (true)
        {
            switch (current)
            {
                case MemberAccessExpressionSyntax memberAccess:
                    current = memberAccess.Expression;
                    continue;
                case InvocationExpressionSyntax invocation:
                    current = invocation.Expression;
                    continue;
                case IdentifierNameSyntax identifier:
                    return identifier.Identifier.Text;
                default:
                    return null;
            }
        }
    }

    /// <summary>
    ///     Checks for .Cell or .Cells member access anywhere in the expression.
    /// </summary>
    private static bool DetectObjectModel(SyntaxNode root)
    {
        return root.DescendantNodes()
            .OfType<MemberAccessExpressionSyntax>()
            .Any(ma =>
            {
                var name = ma.Name.Identifier.Text;
                return name.Equals("Cell", StringComparison.OrdinalIgnoreCase) ||
                       name.Equals("Cells", StringComparison.OrdinalIgnoreCase);
            });
    }

    /// <summary>
    ///     Collects all lambda parameter names declared anywhere in the expression tree.
    /// </summary>
    private static HashSet<string> CollectAllLambdaParameters(SyntaxNode root)
    {
        var result = new HashSet<string>();

        foreach (var lambda in root.DescendantNodes().OfType<SimpleLambdaExpressionSyntax>())
        {
            result.Add(lambda.Parameter.Identifier.Text);
        }

        foreach (var lambda in root.DescendantNodes().OfType<ParenthesizedLambdaExpressionSyntax>())
        {
            foreach (var param in lambda.ParameterList.Parameters)
            {
                result.Add(param.Identifier.Text);
            }
        }

        return result;
    }

    /// <summary>
    ///     Detects free variables — identifiers that are not inputs, lambda parameters,
    ///     C# keywords, known type names, or method names in invocation context.
    /// </summary>
    private static IReadOnlyList<string> DetectFreeVariables(
        SyntaxNode root,
        List<string> inputs,
        HashSet<string> allLambdaParams)
    {
        var inputSet = new HashSet<string>(inputs);
        var freeVars = new HashSet<string>();

        // Collect all locally declared variable names
        var declaredLocals = new HashSet<string>();
        foreach (var varDecl in root.DescendantNodes().OfType<VariableDeclaratorSyntax>())
        {
            declaredLocals.Add(varDecl.Identifier.Text);
        }

        foreach (var identifier in root.DescendantNodes().OfType<IdentifierNameSyntax>())
        {
            var name = identifier.Identifier.Text;

            // Skip if it's a known identifier
            if (inputSet.Contains(name) || allLambdaParams.Contains(name) ||
                IgnoredIdentifiers.Contains(name) || declaredLocals.Contains(name))
            {
                continue;
            }

            // Skip if it starts with __ (internal placeholder)
            if (name.StartsWith("__"))
            {
                continue;
            }

            // Skip if it's a method name being invoked (e.g., Where, Select)
            if (identifier.Parent is MemberAccessExpressionSyntax memberAccess &&
                memberAccess.Name == identifier)
            {
                continue;
            }

            // Skip if it's a member being accessed on something (not standalone)
            if (identifier.Parent is MemberAccessExpressionSyntax parentMa &&
                parentMa.Name == identifier)
            {
                continue;
            }

            // Skip the wrapper class and method artifacts
            if (name == "__Wrapper" || name == "__M")
            {
                continue;
            }

            freeVars.Add(name);
        }

        return freeVars.OrderBy(v => v).ToList();
    }

    /// <summary>
    ///     Detects r["Col Name"] string bracket access patterns.
    /// </summary>
    private static bool DetectStringBracketAccess(SyntaxNode root)
    {
        return root.DescendantNodes()
            .OfType<ElementAccessExpressionSyntax>()
            .Any(ea => ea.ArgumentList.Arguments
                .Any(arg => arg.Expression is LiteralExpressionSyntax lit &&
                            lit.IsKind(SyntaxKind.StringLiteralExpression)));
    }
}
