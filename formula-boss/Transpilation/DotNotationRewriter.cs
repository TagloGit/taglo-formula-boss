using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Roslyn syntax rewriter that converts dot-notation column access on row parameters
///     to bracket (element access) notation using the original column name.
///     e.g. <c>r.Population2025</c> → <c>r["Population 2025"]</c>
/// </summary>
public class DotNotationRewriter : CSharpSyntaxRewriter
{
    private readonly Dictionary<string, string> _columnMapping;
    private readonly HashSet<string> _rowParameterNames;

    /// <param name="columnMapping">Sanitised identifier → original column name.</param>
    /// <param name="rowParameterNames">Lambda parameter names that represent rows (e.g. "r").</param>
    public DotNotationRewriter(
        Dictionary<string, string> columnMapping,
        HashSet<string> rowParameterNames)
    {
        _columnMapping = columnMapping;
        _rowParameterNames = rowParameterNames;
    }

    public override SyntaxNode? VisitMemberAccessExpression(MemberAccessExpressionSyntax node)
    {
        // Check if expression target is a row parameter: r.SomeName
        if (node.Expression is IdentifierNameSyntax identifier &&
            _rowParameterNames.Contains(identifier.Identifier.Text) &&
            _columnMapping.TryGetValue(node.Name.Identifier.Text, out var originalName))
        {
            // Rewrite r.SanitisedName → r["Original Name"]
            return SyntaxFactory.ElementAccessExpression(
                    node.Expression,
                    SyntaxFactory.BracketedArgumentList(
                        SyntaxFactory.SingletonSeparatedList(
                            SyntaxFactory.Argument(
                                SyntaxFactory.LiteralExpression(
                                    SyntaxKind.StringLiteralExpression,
                                    SyntaxFactory.Literal(originalName))))))
                .WithTriviaFrom(node);
        }

        return base.VisitMemberAccessExpression(node);
    }

    /// <summary>
    ///     Rewrites dot-notation column access in an expression string.
    ///     Parses the expression, applies the rewrite, and returns the modified expression string.
    /// </summary>
    public static string Rewrite(
        string expression,
        Dictionary<string, string> columnMapping,
        HashSet<string>? rowParameterNames = null)
    {
        if (columnMapping.Count == 0)
        {
            return expression;
        }

        // Parse the expression in a method context
        var wrappedSource = $"class __W {{ object __M() => {expression}; }}";
        var tree = CSharpSyntaxTree.ParseText(wrappedSource);
        var root = tree.GetRoot();

        // Detect row parameter names from lambdas if not provided
        var rowParams = rowParameterNames ?? DetectRowParameters(root);

        if (rowParams.Count == 0)
        {
            return expression;
        }

        var rewriter = new DotNotationRewriter(columnMapping, rowParams);
        var newRoot = rewriter.Visit(root);

        // Extract the rewritten expression from the wrapper
        var method = newRoot.DescendantNodes().OfType<MethodDeclarationSyntax>().First();
        var rewrittenExpr = method.ExpressionBody?.Expression;

        return rewrittenExpr?.ToFullString() ?? expression;
    }

    /// <summary>
    ///     Detects lambda parameter names that are likely row parameters
    ///     (used in member access or bracket access patterns).
    /// </summary>
    private static HashSet<string> DetectRowParameters(SyntaxNode root)
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
}
