using System.Text.RegularExpressions;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Rewrites a user statement block to emit <see cref="FormulaBoss.Runtime.Tracer" /> calls for
///     debug mode. Inserts <c>Tracer.Begin</c> + an entry snapshot at the top, <c>Tracer.Set</c>
///     after each local declaration / assignment, <c>Tracer.Snapshot("iter", depth, branch)</c> at
///     the end of each loop body, and rewrites <c>return</c> statements to capture the returned
///     value, fire <c>Tracer.Return</c>, and snapshot before returning.
///     Loop depth is tracked via a visit counter. Branch labels are derived from the first
///     statement of each if/else arm (trimmed and truncated to <see cref="BranchLabelMaxLength" />).
/// </summary>
public class DebugInstrumentationRewriter : CSharpSyntaxRewriter
{
    private const string TracerType = "FormulaBoss.Runtime.Tracer";
    private const string ReturnTempName = "__fb_ret";
    private const int BranchLabelMaxLength = 20;

    private static readonly Regex WhitespaceRun = new(@"\s+", RegexOptions.Compiled);

    private string? _branchLabel;
    private int _loopDepth;

    /// <summary>
    ///     Instrument a block. Prepends <c>Tracer.Begin(name, callerAddr)</c> + entry snapshot,
    ///     then recursively injects Set/Snapshot/Return calls inside the block.
    /// </summary>
    /// <param name="block">The user's statement block.</param>
    /// <param name="traceName">Name recorded by <c>Tracer.Begin</c> (usually the UDF method name).</param>
    /// <param name="callerAddressExpression">
    ///     A C# expression (already formatted) that evaluates to the caller cell address string.
    ///     Use <c>"\"\""</c> as a placeholder when the address is not yet wired.
    /// </param>
    public static BlockSyntax Instrument(BlockSyntax block, string traceName, string callerAddressExpression)
    {
        var rewriter = new DebugInstrumentationRewriter();
        var rewritten = (BlockSyntax)rewriter.Visit(block);

        var begin = SyntaxFactory.ParseStatement(
            $"{TracerType}.Begin({Literal(traceName)}, {callerAddressExpression});");
        var entry = SyntaxFactory.ParseStatement(
            $"{TracerType}.Snapshot(\"entry\", 0, null);");

        var statements = new List<StatementSyntax> { begin, entry };
        statements.AddRange(rewritten.Statements);
        return rewritten.WithStatements(SyntaxFactory.List(statements));
    }

    public override SyntaxNode VisitBlock(BlockSyntax node)
    {
        var newStatements = new List<StatementSyntax>();
        foreach (var stmt in node.Statements)
        {
            if (Visit(stmt) is not StatementSyntax visited)
            {
                continue;
            }

            newStatements.Add(visited);

            // Insert Tracer.Set calls for newly-assigned locals.
            switch (stmt)
            {
                case LocalDeclarationStatementSyntax decl:
                    foreach (var v in decl.Declaration.Variables)
                    {
                        newStatements.Add(MakeSet(v.Identifier.Text));
                    }

                    break;
                case ExpressionStatementSyntax { Expression: AssignmentExpressionSyntax assign }
                    when assign.Left is IdentifierNameSyntax id:
                    newStatements.Add(MakeSet(id.Identifier.Text));
                    break;
            }
        }

        return node.WithStatements(SyntaxFactory.List(newStatements));
    }

    public override SyntaxNode VisitForEachStatement(ForEachStatementSyntax node)
    {
        _loopDepth++;
        var body = EnsureBlock((StatementSyntax)Visit(node.Statement));
        var statements = new List<StatementSyntax> { MakeSet(node.Identifier.Text) };
        statements.AddRange(body.Statements);
        statements.Add(MakeSnapshot("iter"));
        var newBody = body.WithStatements(SyntaxFactory.List(statements));
        _loopDepth--;
        return node.WithStatement(newBody);
    }

    public override SyntaxNode VisitForStatement(ForStatementSyntax node)
    {
        _loopDepth++;
        var body = EnsureBlock((StatementSyntax)Visit(node.Statement));
        var statements = new List<StatementSyntax>();
        if (node.Declaration != null)
        {
            foreach (var v in node.Declaration.Variables)
            {
                statements.Add(MakeSet(v.Identifier.Text));
            }
        }

        statements.AddRange(body.Statements);
        statements.Add(MakeSnapshot("iter"));
        var newBody = body.WithStatements(SyntaxFactory.List(statements));
        _loopDepth--;
        return node.WithStatement(newBody);
    }

    public override SyntaxNode VisitWhileStatement(WhileStatementSyntax node)
    {
        _loopDepth++;
        var body = EnsureBlock((StatementSyntax)Visit(node.Statement));
        var statements = new List<StatementSyntax>(body.Statements) { MakeSnapshot("iter") };
        var newBody = body.WithStatements(SyntaxFactory.List(statements));
        _loopDepth--;
        return node.WithStatement(newBody);
    }

    public override SyntaxNode VisitIfStatement(IfStatementSyntax node)
    {
        var savedBranch = _branchLabel;

        _branchLabel = MakeBranchLabel(node.Statement);
        var visitedThen = (StatementSyntax)Visit(node.Statement);

        ElseClauseSyntax? visitedElse = null;
        if (node.Else != null)
        {
            _branchLabel = MakeBranchLabel(node.Else.Statement);
            var visitedElseStmt = (StatementSyntax)Visit(node.Else.Statement);
            visitedElse = node.Else.WithStatement(visitedElseStmt);
        }

        _branchLabel = savedBranch;
        return visitedElse != null
            ? node.WithStatement(visitedThen).WithElse(visitedElse)
            : node.WithStatement(visitedThen);
    }

    public override SyntaxNode VisitReturnStatement(ReturnStatementSyntax node)
    {
        if (node.Expression == null)
        {
            return node;
        }

        var tempDecl = SyntaxFactory.ParseStatement(
            $"var {ReturnTempName} = {node.Expression.ToFullString()};");
        var returnCall = SyntaxFactory.ParseStatement(
            $"{TracerType}.Return({ReturnTempName});");
        var snapshot = MakeSnapshot("return");
        var ret = SyntaxFactory.ParseStatement($"return {ReturnTempName};");

        return SyntaxFactory.Block(tempDecl, returnCall, snapshot, ret);
    }

    private StatementSyntax MakeSet(string name)
        => SyntaxFactory.ParseStatement($"{TracerType}.Set({Literal(name)}, {name});");

    private StatementSyntax MakeSnapshot(string kind)
        => SyntaxFactory.ParseStatement(
            $"{TracerType}.Snapshot({Literal(kind)}, {_loopDepth}, {BranchArg()});");

    private string BranchArg() => _branchLabel == null ? "null" : Literal(_branchLabel);

    private static BlockSyntax EnsureBlock(StatementSyntax stmt)
        => stmt is BlockSyntax b ? b : SyntaxFactory.Block(stmt);

    private static string MakeBranchLabel(StatementSyntax stmt)
    {
        var first = stmt;
        if (stmt is BlockSyntax block)
        {
            first = block.Statements.FirstOrDefault();
        }

        if (first == null)
        {
            return string.Empty;
        }

        var text = WhitespaceRun.Replace(first.ToString().Trim(), " ");
        if (text.Length > BranchLabelMaxLength)
        {
            text = text[..BranchLabelMaxLength];
        }

        return text;
    }

    private static string Literal(string s)
        => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
