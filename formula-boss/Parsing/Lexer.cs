using System.Globalization;
using System.Text;

namespace FormulaBoss.Parsing;

/// <summary>
/// Tokenizes DSL source text into a sequence of tokens.
/// </summary>
public class Lexer
{
    private readonly string _source;
    private int _position;
    private readonly List<Token> _tokens = [];

    public Lexer(string source)
    {
        _source = source;
    }

    /// <summary>
    /// Scans the entire source and returns all tokens.
    /// </summary>
    public List<Token> ScanTokens()
    {
        while (!IsAtEnd())
        {
            ScanToken();
        }

        _tokens.Add(new Token(TokenType.Eof, "", null, _position));
        return _tokens;
    }

    private void ScanToken()
    {
        SkipWhitespace();
        if (IsAtEnd())
        {
            return;
        }

        var start = _position;
        var c = Advance();

        switch (c)
        {
            case '.':
                AddToken(TokenType.Dot, start);
                break;
            case ',':
                AddToken(TokenType.Comma, start);
                break;
            case '+':
                AddToken(TokenType.Plus, start);
                break;
            case '-':
                AddToken(TokenType.Minus, start);
                break;
            case '*':
                AddToken(TokenType.Star, start);
                break;
            case '/':
                AddToken(TokenType.Slash, start);
                break;
            case '(':
                AddToken(TokenType.LeftParen, start);
                break;
            case ')':
                AddToken(TokenType.RightParen, start);
                break;
            case '=':
                if (Match('='))
                {
                    AddToken(TokenType.Equal, start, "==");
                }
                else if (Match('>'))
                {
                    AddToken(TokenType.Arrow, start, "=>");
                }
                else
                {
                    AddErrorToken(start, "Unexpected '='. Did you mean '==' or '=>'?");
                }

                break;
            case '!':
                if (Match('='))
                {
                    AddToken(TokenType.NotEqual, start, "!=");
                }
                else
                {
                    AddToken(TokenType.Not, start);
                }

                break;
            case '>':
                if (Match('='))
                {
                    AddToken(TokenType.GreaterEqual, start, ">=");
                }
                else
                {
                    AddToken(TokenType.Greater, start);
                }

                break;
            case '<':
                if (Match('='))
                {
                    AddToken(TokenType.LessEqual, start, "<=");
                }
                else
                {
                    AddToken(TokenType.Less, start);
                }

                break;
            case '&':
                if (Match('&'))
                {
                    AddToken(TokenType.And, start, "&&");
                }
                else
                {
                    AddErrorToken(start, "Unexpected '&'. Did you mean '&&'?");
                }

                break;
            case '|':
                if (Match('|'))
                {
                    AddToken(TokenType.Or, start, "||");
                }
                else
                {
                    AddErrorToken(start, "Unexpected '|'. Did you mean '||'?");
                }

                break;
            case '"':
                ScanString(start);
                break;
            default:
                if (char.IsDigit(c))
                {
                    ScanNumber(start);
                }
                else if (IsIdentifierStart(c))
                {
                    ScanIdentifier(start);
                }
                else
                {
                    AddErrorToken(start, $"Unexpected character '{c}'");
                }

                break;
        }
    }

    private void ScanString(int start)
    {
        var sb = new StringBuilder();
        while (!IsAtEnd() && Peek() != '"')
        {
            if (Peek() == '\\' && _position + 1 < _source.Length)
            {
                Advance(); // consume backslash
                var escaped = Advance();
                sb.Append(escaped switch
                {
                    'n' => '\n',
                    't' => '\t',
                    'r' => '\r',
                    '"' => '"',
                    '\\' => '\\',
                    _ => escaped
                });
            }
            else
            {
                sb.Append(Advance());
            }
        }

        if (IsAtEnd())
        {
            AddErrorToken(start, "Unterminated string");
            return;
        }

        Advance(); // consume closing quote
        var lexeme = _source[start.._position];
        _tokens.Add(new Token(TokenType.StringLiteral, lexeme, sb.ToString(), start));
    }

    private void ScanNumber(int start)
    {
        while (!IsAtEnd() && char.IsDigit(Peek()))
        {
            Advance();
        }

        // Look for decimal part
        if (!IsAtEnd() && Peek() == '.' && _position + 1 < _source.Length && char.IsDigit(_source[_position + 1]))
        {
            Advance(); // consume '.'
            while (!IsAtEnd() && char.IsDigit(Peek()))
            {
                Advance();
            }
        }

        var lexeme = _source[start.._position];
        var value = double.Parse(lexeme, CultureInfo.InvariantCulture);
        _tokens.Add(new Token(TokenType.Number, lexeme, value, start));
    }

    private void ScanIdentifier(int start)
    {
        while (!IsAtEnd() && IsIdentifierChar(Peek()))
        {
            Advance();
        }

        var lexeme = _source[start.._position];
        _tokens.Add(new Token(TokenType.Identifier, lexeme, null, start));
    }

    private void SkipWhitespace()
    {
        while (!IsAtEnd() && char.IsWhiteSpace(Peek()))
        {
            Advance();
        }
    }

    private bool IsAtEnd() => _position >= _source.Length;

    private char Peek() => _source[_position];

    private char Advance() => _source[_position++];

    private bool Match(char expected)
    {
        if (IsAtEnd() || _source[_position] != expected)
        {
            return false;
        }

        _position++;
        return true;
    }

    private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

    private static bool IsIdentifierChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    private void AddToken(TokenType type, int start, string? lexeme = null)
    {
        lexeme ??= _source[start.._position];
        _tokens.Add(new Token(type, lexeme, null, start));
    }

    private void AddErrorToken(int start, string message)
    {
        _tokens.Add(new Token(TokenType.Error, message, null, start));
    }
}
