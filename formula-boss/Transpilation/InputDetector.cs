using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Result of input detection on a Formula Boss expression.
/// </summary>
/// <param name="Inputs">Ordered list of input parameter names from the expression.</param>
/// <param name="Body">The expression body (everything after the lambda arrow, or the whole expression for sugar syntax).</param>
/// <param name="IsExplicitLambda">True if the expression uses explicit lambda syntax: (a, b) => ...</param>
/// <param name="RequiresObjectModel">True if .Cell or .Cells is used, requiring IsMacroType.</param>
/// <param name="FreeVariables">
///     Identifiers used in the body that are not inputs or lambda parameters.
///     These become additional UDF parameters for LET variable capture.
/// </param>
/// <param name="IsStatementBody">True if the body is a statement block { ... } rather than an expression.</param>
public record InputDetectionResult(
    IReadOnlyList<string> Inputs,
    string Body,
    bool IsExplicitLambda,
    bool RequiresObjectModel,
    IReadOnlyList<string> FreeVariables,
    bool IsStatementBody);

/// <summary>
///     Uses Roslyn to detect inputs, free variables, and cell usage in a Formula Boss expression.
/// </summary>
public static class InputDetector
{
    /// <summary>
    ///     Analyses a Formula Boss expression and extracts input information.
    /// </summary>
    /// <param name="expression">The expression (without backticks).</param>
    /// <param name="knownLetVariables">
    ///     Names of LET variables available in scope. Used to identify free variables
    ///     that should become additional UDF parameters.
    /// </param>
    public static InputDetectionResult Detect(string expression, IReadOnlySet<string>? knownLetVariables = null)
    {
        var trimmed = expression.Trim();

        // Try to parse as a lambda first: (a, b) => body or a => body
        var lambdaResult = TryParseLambda(trimmed);
        if (lambdaResult != null)
        {
            return lambdaResult.Value.Finalize(knownLetVariables);
        }

        // Single-input sugar: identifier.Method(...) or identifier[...]
        // The first identifier before the first '.' is the input
        return ParseSugarSyntax(trimmed, knownLetVariables);
    }

    private static PartialResult? TryParseLambda(string expression)
    {
        // Wrap in a method to make it a valid C# statement for Roslyn
        var wrapper = $"var __x__ = {expression};";
        var tree = CSharpSyntaxTree.ParseText(wrapper, new CSharpParseOptions(LanguageVersion.CSharp10));
        var root = tree.GetCompilationUnitRoot();

        // Find the TOP-LEVEL lambda expression (the RHS of var __x__ = ...)
        // We only want lambdas that ARE the entire expression, not nested ones in method arguments
        var varDecl = root.DescendantNodes().OfType<VariableDeclaratorSyntax>().FirstOrDefault();
        var initExpr = varDecl?.Initializer?.Value;
        var lambda = initExpr as LambdaExpressionSyntax;
        if (lambda == null)
        {
            return null;
        }

        var inputs = new List<string>();

        switch (lambda)
        {
            case ParenthesizedLambdaExpressionSyntax parenLambda:
                inputs.AddRange(parenLambda.ParameterList.Parameters.Select(p => p.Identifier.Text));
                break;
            case SimpleLambdaExpressionSyntax simpleLambda:
                inputs.Add(simpleLambda.Parameter.Identifier.Text);
                break;
            default:
                return null;
        }

        if (inputs.Count == 0)
        {
            return null;
        }

        // Extract body text from original expression
        var body = lambda.Body;
        var bodySpan = body.Span;
        // Map back to original expression: wrapper has "var __x__ = " prefix (12 chars)
        const int prefixLen = 12; // "var __x__ = "
        var bodyStart = bodySpan.Start - prefixLen;
        var bodyLength = bodySpan.Length;

        string bodyText;
        bool isStatementBody;
        if (bodyStart >= 0 && bodyStart + bodyLength <= expression.Length)
        {
            bodyText = expression.Substring(bodyStart, bodyLength);
            isStatementBody = body is BlockSyntax;
        }
        else
        {
            // Fallback: use Roslyn's text
            bodyText = body.ToFullString().Trim();
            isStatementBody = body is BlockSyntax;
        }

        // Detect .Cell/.Cells usage and collect free variables from body
        var requiresObjectModel = DetectCellUsage(body);
        var freeVars = CollectFreeVariables(body, inputs);

        return new PartialResult(inputs, bodyText, true, requiresObjectModel, freeVars, isStatementBody);
    }

    private static InputDetectionResult ParseSugarSyntax(string expression, IReadOnlySet<string>? knownLetVariables)
    {
        // Parse as an expression to extract the primary input
        var wrapper = $"var __x__ = {expression};";
        var tree = CSharpSyntaxTree.ParseText(wrapper, new CSharpParseOptions(LanguageVersion.CSharp10));
        var root = tree.GetCompilationUnitRoot();

        // Find the outermost expression (the RHS of the assignment)
        var varDecl = root.DescendantNodes().OfType<VariableDeclaratorSyntax>().FirstOrDefault();
        var expr = varDecl?.Initializer?.Value;

        string primaryInput;
        if (expr != null)
        {
            primaryInput = ExtractPrimaryInput(expr);
        }
        else
        {
            // Fallback: take everything before the first dot
            var dotIdx = expression.IndexOf('.');
            primaryInput = dotIdx > 0 ? expression[..dotIdx].Trim() : expression.Trim();
        }

        var inputs = new List<string> { primaryInput };

        // Detect .Cell/.Cells and free variables from the full expression
        bool requiresObjectModel;
        IReadOnlyList<string> freeVars;
        if (expr != null)
        {
            requiresObjectModel = DetectCellUsage(expr);
            freeVars = CollectFreeVariables(expr, inputs);
        }
        else
        {
            requiresObjectModel = DetectCellUsageByText(expression);
            freeVars = Array.Empty<string>();
        }

        // Filter free variables to only those known as LET variables
        var filteredFreeVars = FilterToKnownVariables(freeVars, knownLetVariables);

        return new InputDetectionResult(inputs, expression, false, requiresObjectModel, filteredFreeVars, false);
    }

    /// <summary>
    ///     Extracts the primary (leftmost) input identifier from an expression tree.
    ///     For member access chains like tblCountries.Rows.Where(...), returns "tblCountries".
    /// </summary>
    private static string ExtractPrimaryInput(ExpressionSyntax expr)
    {
        // Walk down the left side of member access / invocation chains
        var current = expr;
        while (true)
        {
            switch (current)
            {
                case InvocationExpressionSyntax invocation:
                    current = invocation.Expression;
                    continue;
                case MemberAccessExpressionSyntax memberAccess:
                    current = memberAccess.Expression;
                    continue;
                case ElementAccessExpressionSyntax elementAccess:
                    current = elementAccess.Expression;
                    continue;
                case IdentifierNameSyntax identifier:
                    return identifier.Identifier.Text;
                default:
                    return current.ToString();
            }
        }
    }

    /// <summary>
    ///     Detects usage of .Cell or .Cells in a syntax node.
    /// </summary>
    private static bool DetectCellUsage(SyntaxNode node)
    {
        return node.DescendantNodesAndSelf()
            .OfType<MemberAccessExpressionSyntax>()
            .Any(m => m.Name.Identifier.Text is "Cell" or "Cells");
    }

    private static bool DetectCellUsageByText(string expression) =>
        // Simple text-based fallback
        expression.Contains(".Cell") || expression.Contains(".Cells");

    /// <summary>
    ///     Collects identifiers used in the node that are not in the bound set
    ///     (inputs or lambda parameters). These are potential free variables.
    /// </summary>
    private static IReadOnlyList<string> CollectFreeVariables(SyntaxNode node, IReadOnlyList<string> boundNames)
    {
        var bound = new HashSet<string>(boundNames);
        var freeVars = new LinkedHashSet();

        // Also collect all lambda parameter names as bound
        foreach (var lambda in node.DescendantNodesAndSelf().OfType<LambdaExpressionSyntax>())
        {
            switch (lambda)
            {
                case SimpleLambdaExpressionSyntax simple:
                    bound.Add(simple.Parameter.Identifier.Text);
                    break;
                case ParenthesizedLambdaExpressionSyntax paren:
                    foreach (var p in paren.ParameterList.Parameters)
                    {
                        bound.Add(p.Identifier.Text);
                    }

                    break;
            }
        }

        // Find all identifier usages that aren't:
        // - part of a member access (right side)
        // - bound variables
        // - well-known C# type/method names
        foreach (var identifier in node.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>())
        {
            var name = identifier.Identifier.Text;

            // Skip if bound
            if (bound.Contains(name))
            {
                continue;
            }

            // Skip if it's the right side of a member access (e.g., .Where, .Rows)
            if (identifier.Parent is MemberAccessExpressionSyntax memberAccess &&
                memberAccess.Name == identifier)
            {
                continue;
            }

            // Skip well-known C#/system identifiers
            if (IsWellKnownIdentifier(name))
            {
                continue;
            }

            freeVars.Add(name);
        }

        return freeVars.ToList();
    }

    private static bool IsWellKnownIdentifier(string name)
    {
        return name is "var" or "true" or "false" or "null" or "new"
            or "string" or "int" or "double" or "bool" or "object" or "decimal" or "float" or "long"
            or "Math" or "Convert" or "Console" or "String" or "Int32" or "Double" or "Boolean"
            or "Regex" or "DateTime" or "TimeSpan" or "Guid" or "Enumerable"
            or "Rows" or "Cells" or "Cell" or "Value" or "ToResult";
    }

    private static IReadOnlyList<string> FilterToKnownVariables(
        IReadOnlyList<string> freeVars,
        IReadOnlySet<string>? knownLetVariables)
    {
        if (knownLetVariables == null || knownLetVariables.Count == 0)
        {
            return freeVars;
        }

        return freeVars.Where(v => knownLetVariables.Contains(v)).ToList();
    }

    /// <summary>
    ///     Intermediate result before finalizing with known LET variables.
    /// </summary>
    private readonly record struct PartialResult(
        IReadOnlyList<string> Inputs,
        string Body,
        bool IsExplicitLambda,
        bool RequiresObjectModel,
        IReadOnlyList<string> FreeVariables,
        bool IsStatementBody)
    {
        public InputDetectionResult Finalize(IReadOnlySet<string>? knownLetVariables)
        {
            var filteredFreeVars = FilterToKnownVariables(FreeVariables, knownLetVariables);
            return new InputDetectionResult(Inputs, Body, IsExplicitLambda, RequiresObjectModel,
                filteredFreeVars, IsStatementBody);
        }
    }

    /// <summary>
    ///     Insertion-order-preserving set of strings.
    /// </summary>
    private sealed class LinkedHashSet
    {
        private readonly List<string> _list = [];
        private readonly HashSet<string> _set = [];

        public void Add(string item)
        {
            if (_set.Add(item))
            {
                _list.Add(item);
            }
        }

        public List<string> ToList() => [.. _list];
    }
}
