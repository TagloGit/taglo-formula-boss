using FormulaBoss.Parsing;

using Xunit;

namespace FormulaBoss.Tests;

public class ParserTests
{
    #region Lexer Tests

    [Fact]
    public void Lexer_SimpleIdentifier()
    {
        var lexer = new Lexer("data");
        var tokens = lexer.ScanTokens();

        Assert.Equal(2, tokens.Count);
        Assert.Equal(TokenType.Identifier, tokens[0].Type);
        Assert.Equal("data", tokens[0].Lexeme);
        Assert.Equal(TokenType.Eof, tokens[1].Type);
    }

    [Fact]
    public void Lexer_Number()
    {
        var lexer = new Lexer("42.5");
        var tokens = lexer.ScanTokens();

        Assert.Equal(2, tokens.Count);
        Assert.Equal(TokenType.Number, tokens[0].Type);
        Assert.Equal(42.5, tokens[0].Literal);
    }

    [Fact]
    public void Lexer_String()
    {
        var lexer = new Lexer("\"hello\"");
        var tokens = lexer.ScanTokens();

        Assert.Equal(2, tokens.Count);
        Assert.Equal(TokenType.StringLiteral, tokens[0].Type);
        Assert.Equal("hello", tokens[0].Literal);
    }

    [Fact]
    public void Lexer_Operators()
    {
        var lexer = new Lexer("== != > < >= <= && || =>");
        var tokens = lexer.ScanTokens();

        Assert.Equal(TokenType.Equal, tokens[0].Type);
        Assert.Equal(TokenType.NotEqual, tokens[1].Type);
        Assert.Equal(TokenType.Greater, tokens[2].Type);
        Assert.Equal(TokenType.Less, tokens[3].Type);
        Assert.Equal(TokenType.GreaterEqual, tokens[4].Type);
        Assert.Equal(TokenType.LessEqual, tokens[5].Type);
        Assert.Equal(TokenType.And, tokens[6].Type);
        Assert.Equal(TokenType.Or, tokens[7].Type);
        Assert.Equal(TokenType.Arrow, tokens[8].Type);
    }

    [Fact]
    public void Lexer_MethodChain()
    {
        var lexer = new Lexer("data.cells.where(c => c.color == 6)");
        var tokens = lexer.ScanTokens();

        var expectedTypes = new[]
        {
            TokenType.Identifier, // data
            TokenType.Dot,
            TokenType.Identifier, // cells
            TokenType.Dot,
            TokenType.Identifier, // where
            TokenType.LeftParen,
            TokenType.Identifier, // c
            TokenType.Arrow,
            TokenType.Identifier, // c
            TokenType.Dot,
            TokenType.Identifier, // color
            TokenType.Equal,
            TokenType.Number, // 6
            TokenType.RightParen,
            TokenType.Eof
        };

        Assert.Equal(expectedTypes.Length, tokens.Count);
        for (var i = 0; i < expectedTypes.Length; i++)
        {
            Assert.Equal(expectedTypes[i], tokens[i].Type);
        }
    }

    #endregion

    #region Parser Tests - Literals and Identifiers

    [Fact]
    public void Parser_Identifier()
    {
        var expr = Parse("data");

        var ident = Assert.IsType<IdentifierExpr>(expr);
        Assert.Equal("data", ident.Name);
    }

    [Fact]
    public void Parser_Number()
    {
        var expr = Parse("42");

        var num = Assert.IsType<NumberLiteral>(expr);
        Assert.Equal(42, num.Value);
    }

    [Fact]
    public void Parser_String()
    {
        var expr = Parse("\"hello\"");

        var str = Assert.IsType<StringLiteral>(expr);
        Assert.Equal("hello", str.Value);
    }

    #endregion

    #region Parser Tests - Member Access

    [Fact]
    public void Parser_MemberAccess()
    {
        var expr = Parse("data.cells");

        var member = Assert.IsType<MemberAccess>(expr);
        Assert.Equal("cells", member.Member);

        var obj = Assert.IsType<IdentifierExpr>(member.Target);
        Assert.Equal("data", obj.Name);
    }

    [Fact]
    public void Parser_ChainedMemberAccess()
    {
        var expr = Parse("a.b.c");

        var c = Assert.IsType<MemberAccess>(expr);
        Assert.Equal("c", c.Member);

        var b = Assert.IsType<MemberAccess>(c.Target);
        Assert.Equal("b", b.Member);

        var a = Assert.IsType<IdentifierExpr>(b.Target);
        Assert.Equal("a", a.Name);
    }

    #endregion

    #region Parser Tests - Method Calls

    [Fact]
    public void Parser_MethodCallNoArgs()
    {
        var expr = Parse("data.toArray()");

        var call = Assert.IsType<MethodCall>(expr);
        Assert.Equal("toArray", call.Method);
        Assert.Empty(call.Arguments);
    }

    [Fact]
    public void Parser_MethodCallWithArgs()
    {
        var expr = Parse("list.take(5)");

        var call = Assert.IsType<MethodCall>(expr);
        Assert.Equal("take", call.Method);
        Assert.Single(call.Arguments);

        var arg = Assert.IsType<NumberLiteral>(call.Arguments[0]);
        Assert.Equal(5, arg.Value);
    }

    [Fact]
    public void Parser_MethodChain()
    {
        var expr = Parse("data.cells.toArray()");

        var toArray = Assert.IsType<MethodCall>(expr);
        Assert.Equal("toArray", toArray.Method);

        var cells = Assert.IsType<MemberAccess>(toArray.Target);
        Assert.Equal("cells", cells.Member);
    }

    #endregion

    #region Parser Tests - Lambdas

    [Fact]
    public void Parser_LambdaInMethodCall()
    {
        var expr = Parse("data.where(c => c.value)");

        var call = Assert.IsType<MethodCall>(expr);
        Assert.Equal("where", call.Method);
        Assert.Single(call.Arguments);

        var lambda = Assert.IsType<LambdaExpr>(call.Arguments[0]);
        Assert.Equal("c", lambda.Parameter);

        var body = Assert.IsType<MemberAccess>(lambda.Body);
        Assert.Equal("value", body.Member);
    }

    [Fact]
    public void Parser_LambdaWithComparison()
    {
        var expr = Parse("data.where(c => c.color == 6)");

        var call = Assert.IsType<MethodCall>(expr);
        var lambda = Assert.IsType<LambdaExpr>(call.Arguments[0]);

        var body = Assert.IsType<BinaryExpr>(lambda.Body);
        Assert.Equal("==", body.Operator);

        var left = Assert.IsType<MemberAccess>(body.Left);
        Assert.Equal("color", left.Member);

        var right = Assert.IsType<NumberLiteral>(body.Right);
        Assert.Equal(6, right.Value);
    }

    #endregion

    #region Parser Tests - Operators

    [Fact]
    public void Parser_ArithmeticPrecedence()
    {
        // a + b * c should parse as a + (b * c)
        var expr = Parse("a + b * c");

        var add = Assert.IsType<BinaryExpr>(expr);
        Assert.Equal("+", add.Operator);

        Assert.IsType<IdentifierExpr>(add.Left);

        var mul = Assert.IsType<BinaryExpr>(add.Right);
        Assert.Equal("*", mul.Operator);
    }

    [Fact]
    public void Parser_ComparisonOperators()
    {
        var expr = Parse("a > b");

        var cmp = Assert.IsType<BinaryExpr>(expr);
        Assert.Equal(">", cmp.Operator);
    }

    [Fact]
    public void Parser_LogicalOperators()
    {
        // a && b || c should parse as (a && b) || c
        var expr = Parse("a && b || c");

        var or = Assert.IsType<BinaryExpr>(expr);
        Assert.Equal("||", or.Operator);

        var and = Assert.IsType<BinaryExpr>(or.Left);
        Assert.Equal("&&", and.Operator);
    }

    [Fact]
    public void Parser_UnaryNot()
    {
        var expr = Parse("!flag");

        var unary = Assert.IsType<UnaryExpr>(expr);
        Assert.Equal("!", unary.Operator);

        var ident = Assert.IsType<IdentifierExpr>(unary.Operand);
        Assert.Equal("flag", ident.Name);
    }

    [Fact]
    public void Parser_UnaryMinus()
    {
        var expr = Parse("-5");

        var unary = Assert.IsType<UnaryExpr>(expr);
        Assert.Equal("-", unary.Operator);

        var num = Assert.IsType<NumberLiteral>(unary.Operand);
        Assert.Equal(5, num.Value);
    }

    #endregion

    #region Parser Tests - Grouping

    [Fact]
    public void Parser_Grouping()
    {
        var expr = Parse("(a + b) * c");

        var mul = Assert.IsType<BinaryExpr>(expr);
        Assert.Equal("*", mul.Operator);

        var group = Assert.IsType<GroupingExpr>(mul.Left);
        var add = Assert.IsType<BinaryExpr>(group.Inner);
        Assert.Equal("+", add.Operator);
    }

    #endregion

    #region Parser Tests - Full DSL Expressions

    [Fact]
    public void Parser_FullDslExpression()
    {
        var expr = Parse("data.cells.where(c => c.color == 6).select(c => c.value).toArray()");

        // Should parse to: toArray(select(where(data.cells, lambda), lambda))
        var toArray = Assert.IsType<MethodCall>(expr);
        Assert.Equal("toArray", toArray.Method);

        var select = Assert.IsType<MethodCall>(toArray.Target);
        Assert.Equal("select", select.Method);

        var where = Assert.IsType<MethodCall>(select.Target);
        Assert.Equal("where", where.Method);

        var cells = Assert.IsType<MemberAccess>(where.Target);
        Assert.Equal("cells", cells.Member);

        var data = Assert.IsType<IdentifierExpr>(cells.Target);
        Assert.Equal("data", data.Name);
    }

    #endregion

    #region Parser Tests - Error Handling

    [Fact]
    public void Parser_ErrorOnInvalidInput()
    {
        var lexer = new Lexer("data.");
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);

        var result = parser.Parse();

        Assert.Null(result);
        Assert.NotEmpty(parser.Errors);
    }

    [Fact]
    public void Parser_ErrorOnUnmatchedParen()
    {
        var lexer = new Lexer("(a + b");
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);

        var result = parser.Parse();

        Assert.Null(result);
        Assert.NotEmpty(parser.Errors);
    }

    #endregion

    private static Expression? Parse(string source)
    {
        var lexer = new Lexer(source);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);
        return parser.Parse();
    }
}
