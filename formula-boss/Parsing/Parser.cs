namespace FormulaBoss.Parsing;

/// <summary>
/// Parses a list of tokens into an AST.
/// </summary>
public class Parser
{
    private readonly List<Token> _tokens;
    private int _current;
    private readonly List<string> _errors = [];

    public Parser(List<Token> tokens)
    {
        _tokens = tokens;
    }

    /// <summary>
    /// Gets any parsing errors that occurred.
    /// </summary>
    public IReadOnlyList<string> Errors => _errors;

    /// <summary>
    /// Parses the tokens into an expression AST.
    /// </summary>
    /// <returns>The parsed expression, or null if parsing failed.</returns>
    public Expression? Parse()
    {
        try
        {
            var expr = ParseExpression();
            if (!IsAtEnd())
            {
                Error($"Unexpected token '{Current().Lexeme}' at position {Current().Position}");
            }

            return _errors.Count == 0 ? expr : null;
        }
        catch (ParseException)
        {
            return null;
        }
    }

    // Expression parsing with operator precedence (lowest to highest)

    private Expression ParseExpression() => ParseOr();

    private Expression ParseOr()
    {
        var expr = ParseAnd();

        while (Match(TokenType.Or))
        {
            var right = ParseAnd();
            expr = new BinaryExpr(expr, "||", right);
        }

        return expr;
    }

    private Expression ParseAnd()
    {
        var expr = ParseEquality();

        while (Match(TokenType.And))
        {
            var right = ParseEquality();
            expr = new BinaryExpr(expr, "&&", right);
        }

        return expr;
    }

    private Expression ParseEquality()
    {
        var expr = ParseComparison();

        while (Match(TokenType.Equal, TokenType.NotEqual))
        {
            var op = Previous().Type == TokenType.Equal ? "==" : "!=";
            var right = ParseComparison();
            expr = new BinaryExpr(expr, op, right);
        }

        return expr;
    }

    private Expression ParseComparison()
    {
        var expr = ParseTerm();

        while (Match(TokenType.Greater, TokenType.GreaterEqual, TokenType.Less, TokenType.LessEqual))
        {
            var op = Previous().Type switch
            {
                TokenType.Greater => ">",
                TokenType.GreaterEqual => ">=",
                TokenType.Less => "<",
                TokenType.LessEqual => "<=",
                _ => throw new InvalidOperationException()
            };
            var right = ParseTerm();
            expr = new BinaryExpr(expr, op, right);
        }

        return expr;
    }

    private Expression ParseTerm()
    {
        var expr = ParseFactor();

        while (Match(TokenType.Plus, TokenType.Minus))
        {
            var op = Previous().Type == TokenType.Plus ? "+" : "-";
            var right = ParseFactor();
            expr = new BinaryExpr(expr, op, right);
        }

        return expr;
    }

    private Expression ParseFactor()
    {
        var expr = ParseUnary();

        while (Match(TokenType.Star, TokenType.Slash))
        {
            var op = Previous().Type == TokenType.Star ? "*" : "/";
            var right = ParseUnary();
            expr = new BinaryExpr(expr, op, right);
        }

        return expr;
    }

    private Expression ParseUnary()
    {
        if (Match(TokenType.Not, TokenType.Minus))
        {
            var op = Previous().Type == TokenType.Not ? "!" : "-";
            var operand = ParseUnary();
            return new UnaryExpr(op, operand);
        }

        return ParsePostfix();
    }

    private Expression ParsePostfix()
    {
        var expr = ParsePrimary();

        while (true)
        {
            if (Match(TokenType.Dot))
            {
                var name = Consume(TokenType.Identifier, "Expected identifier after '.'");
                if (Match(TokenType.LeftParen))
                {
                    // Method call
                    var args = ParseArguments();
                    Consume(TokenType.RightParen, "Expected ')' after arguments");
                    expr = new MethodCall(expr, name.Lexeme, args);
                }
                else
                {
                    // Member access
                    expr = new MemberAccess(expr, name.Lexeme);
                }
            }
            else
            {
                break;
            }
        }

        return expr;
    }

    private List<Expression> ParseArguments()
    {
        var args = new List<Expression>();

        if (!Check(TokenType.RightParen))
        {
            do
            {
                args.Add(ParseArgumentExpression());
            }
            while (Match(TokenType.Comma));
        }

        return args;
    }

    private Expression ParseArgumentExpression()
    {
        // Check for lambda: identifier => expression
        if (Check(TokenType.Identifier) && CheckNext(TokenType.Arrow))
        {
            var param = Advance();
            Advance(); // consume =>
            var body = ParseExpression();
            return new LambdaExpr(param.Lexeme, body);
        }

        return ParseExpression();
    }

    private Expression ParsePrimary()
    {
        if (Match(TokenType.Number))
        {
            return new NumberLiteral((double)Previous().Literal!);
        }

        if (Match(TokenType.StringLiteral))
        {
            return new StringLiteral((string)Previous().Literal!);
        }

        if (Match(TokenType.Identifier))
        {
            return new IdentifierExpr(Previous().Lexeme);
        }

        if (Match(TokenType.LeftParen))
        {
            var expr = ParseExpression();
            Consume(TokenType.RightParen, "Expected ')' after expression");
            return new GroupingExpr(expr);
        }

        if (Check(TokenType.Error))
        {
            var errorToken = Advance();
            throw Error($"Lexer error: {errorToken.Lexeme}");
        }

        throw Error($"Expected expression at position {Current().Position}, got '{Current().Lexeme}'");
    }

    // Helper methods

    private bool Match(params TokenType[] types)
    {
        foreach (var type in types)
        {
            if (Check(type))
            {
                Advance();
                return true;
            }
        }

        return false;
    }

    private bool Check(TokenType type) => !IsAtEnd() && Current().Type == type;

    private bool CheckNext(TokenType type) =>
        _current + 1 < _tokens.Count && _tokens[_current + 1].Type == type;

    private Token Advance()
    {
        if (!IsAtEnd())
        {
            _current++;
        }

        return Previous();
    }

    private bool IsAtEnd() => Current().Type == TokenType.Eof;

    private Token Current() => _tokens[_current];

    private Token Previous() => _tokens[_current - 1];

    private Token Consume(TokenType type, string message)
    {
        if (Check(type))
        {
            return Advance();
        }

        throw Error(message);
    }

    private ParseException Error(string message)
    {
        _errors.Add(message);
        return new ParseException(message);
    }

    private sealed class ParseException : Exception
    {
        public ParseException(string message) : base(message)
        {
        }
    }
}
