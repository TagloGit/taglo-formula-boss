using System.Text.RegularExpressions;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Result of analyzing a user expression for parameters, object model usage, and header variables.
/// </summary>
/// <param name="Parameters">Flat list of all detected parameter names (free variables), in sorted order.</param>
/// <param name="RequiresObjectModel">True if .Cell or .Cells is used (needs IsMacroType).</param>
/// <param name="HeaderVariables">Parameter names that need header extraction (use r["Col"] syntax).</param>
/// <param name="NormalizedExpression">Expression with range refs replaced by valid C# identifiers.</param>
/// <param name="RangeRefMap">Maps placeholder identifiers back to original range refs (e.g., "__range_A1_B10" → "A1:B10").</param>
public record DetectionResult(
    IReadOnlyList<string> Parameters,
    bool RequiresObjectModel,
    IReadOnlySet<string> HeaderVariables,
    string NormalizedExpression,
    IReadOnlyDictionary<string, string> RangeRefMap,
    bool IsStatementBlock = false);

/// <summary>
///     Analyzes user expressions using Roslyn syntax trees to detect parameters,
///     object model usage, header variables, and range references.
///     All inputs are detected via free variable analysis — no explicit lambda syntax.
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
        "true",
        "false",
        "null",
        "new",
        "typeof",
        "nameof",
        "default",
        "is",
        "as",
        "in",
        "var",
        "return",
        "if",
        "else",
        "for",
        "foreach",
        "while",
        "do",
        "switch",
        "case",
        "break",
        "continue",
        "throw",
        "try",
        "catch",
        "finally",
        // Common type names
        "string",
        "int",
        "double",
        "float",
        "bool",
        "object",
        "decimal",
        "long",
        "char",
        "byte",
        "Convert",
        "Math",
        "String",
        "Int32",
        "Double",
        "Boolean",
        "Enumerable",
        "System",
        "Linq",
        // Runtime type names (these resolve at runtime, not as free vars)
        "ExcelValue",
        "ExcelArray",
        "ExcelTable",
        "ExcelScalar",
        "Row",
        "ColumnValue",
        "Cell",
        "Interior",
        "CellFont",
        "RangeOrigin",
        "ResultConverter",
        "IExcelRange",
        "RuntimeBridge"
    };

    /// <summary>
    ///     Analyzes a user expression and returns detection results.
    ///     All inputs are detected as free variables — no explicit outer lambda syntax.
    /// </summary>
    public DetectionResult Detect(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            throw new TranspileException("Expression is empty");
        }

        // Step 1: Pre-process range references
        var (normalized, rangeRefMap) = PreprocessRangeRefs(expression);

        // Step 2: Detect statement blocks and parse with Roslyn
        var isStatementBlock = IsStatementBlock(normalized);
        var wrappedSource = isStatementBlock
            ? $"class __Wrapper {{ object __M() {normalized} }}"
            : $"class __Wrapper {{ object __M() => {normalized}; }}";
        var tree = CSharpSyntaxTree.ParseText(wrappedSource);
        var root = tree.GetRoot();

        // Validate parse succeeded
        var methodDecl = root.DescendantNodes().OfType<MethodDeclarationSyntax>().First();
        if (!isStatementBlock && methodDecl.ExpressionBody?.Expression == null)
        {
            throw new TranspileException($"Failed to parse expression: {expression}");
        }

        if (isStatementBlock && methodDecl.Body == null)
        {
            throw new TranspileException($"Failed to parse statement block: {expression}");
        }

        // Step 3: Detect object model usage (.Cell, .Cells)
        var requiresObjectModel = DetectObjectModel(root);

        // Step 4: Detect all free variables as parameters
        var allLambdaParams = CollectAllLambdaParameters(root);
        var parameters = DetectFreeVariables(root, allLambdaParams);

        // Step 5: Detect per-variable header access
        var headerVariables = DetectHeaderVariables(root, [.. parameters]);

        return new DetectionResult(
            parameters,
            requiresObjectModel,
            headerVariables,
            normalized,
            rangeRefMap,
            isStatementBlock);
    }

    /// <summary>
    ///     Detects whether a normalized expression is a statement block (starts with '{' and contains 'return').
    /// </summary>
    internal static bool IsStatementBlock(string normalizedExpression)
    {
        var trimmed = normalizedExpression.TrimStart();
        return trimmed.StartsWith('{') && trimmed.Contains("return");
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
    ///     Detects free variables — identifiers that are not lambda parameters,
    ///     C# keywords, known type names, or method names in invocation context.
    ///     All free variables become UDF parameters.
    /// </summary>
    private static IReadOnlyList<string> DetectFreeVariables(
        SyntaxNode root,
        HashSet<string> allLambdaParams)
    {
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
            if (allLambdaParams.Contains(name) ||
                IgnoredIdentifiers.Contains(name) || declaredLocals.Contains(name))
            {
                continue;
            }

            // Skip if it starts with __ (internal placeholder) but NOT range ref placeholders
            if (name.StartsWith("__") && !name.StartsWith("__range_"))
            {
                continue;
            }

            // Skip if it's a method name being invoked (e.g., Where, Select)
            if (identifier.Parent is MemberAccessExpressionSyntax memberAccess &&
                memberAccess.Name == identifier)
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
    ///     Detects which parameters need header extraction by tracing string bracket
    ///     accesses (r["Col"]) back to their root parameter identifier.
    /// </summary>
    private static IReadOnlySet<string> DetectHeaderVariables(
        SyntaxNode root,
        HashSet<string> parameters)
    {
        var result = new HashSet<string>();

        // Find all element accesses with string literal args
        var stringAccesses = root.DescendantNodes()
            .OfType<ElementAccessExpressionSyntax>()
            .Where(ea => ea.ArgumentList.Arguments
                .Any(arg => arg.Expression is LiteralExpressionSyntax lit &&
                            lit.IsKind(SyntaxKind.StringLiteralExpression)))
            .ToList();

        if (stringAccesses.Count == 0)
        {
            return result;
        }

        foreach (var access in stringAccesses)
        {
            // The receiver of the bracket access (e.g., `r` in `r["Price"]`) is typically
            // a lambda param. Find the enclosing lambda, then trace the invocation chain
            // back to the root parameter.
            var enclosingLambda = access.Ancestors()
                .FirstOrDefault(n => n is SimpleLambdaExpressionSyntax or ParenthesizedLambdaExpressionSyntax);

            if (enclosingLambda != null)
            {
                // Find the invocation containing this lambda (e.g., `.Where(r => r["Price"] > 5)`)
                var invocation = enclosingLambda.Ancestors()
                    .OfType<InvocationExpressionSyntax>().FirstOrDefault();

                if (invocation != null)
                {
                    // Trace the method call chain back to root identifier
                    var rootId = TraceToRootIdentifier(invocation.Expression);
                    if (rootId != null && parameters.Contains(rootId))
                    {
                        result.Add(rootId);
                    }
                }
            }
            else
            {
                // Direct access on a parameter: `tbl["Col"]`
                if (access.Expression is IdentifierNameSyntax id && parameters.Contains(id.Identifier.Text))
                {
                    result.Add(id.Identifier.Text);
                }
            }
        }

        return result;
    }

    /// <summary>
    ///     Traces a member access / invocation chain to its root identifier.
    /// </summary>
    private static string? TraceToRootIdentifier(ExpressionSyntax expr)
    {
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
}
