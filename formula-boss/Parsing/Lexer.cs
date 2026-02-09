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
    /// Gets the source string being tokenized.
    /// </summary>
    public string Source => _source;

    /// <summary>
    /// Gets the current position in the source string.
    /// </summary>
    public int Position => _position;

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
            case '[':
                AddToken(TokenType.LeftBracket, start);
                break;
            case ']':
                AddToken(TokenType.RightBracket, start);
                break;
            case '{':
                AddToken(TokenType.LeftBrace, start);
                break;
            case '}':
                AddToken(TokenType.RightBrace, start);
                break;
            case ';':
                AddToken(TokenType.Semicolon, start);
                break;
            case ':':
                AddToken(TokenType.Colon, start);
                break;
            case '@':
                AddToken(TokenType.At, start);
                break;
            case '?':
                if (Match('?'))
                {
                    AddToken(TokenType.QuestionQuestion, start, "??");
                }
                else
                {
                    AddToken(TokenType.Question, start);
                }

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
                    // Single = is valid inside statement lambdas (C# assignment)
                    AddToken(TokenType.Assign, start);
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
            case '\'':
                ScanCharLiteral(start);
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

    private void ScanCharLiteral(int start)
    {
        // Handle escape sequences
        char value;
        if (!IsAtEnd() && Peek() == '\\')
        {
            Advance(); // consume backslash
            if (IsAtEnd())
            {
                AddErrorToken(start, "Unterminated character literal");
                return;
            }
            var escaped = Advance();
            value = escaped switch
            {
                'n' => '\n',
                't' => '\t',
                'r' => '\r',
                '\'' => '\'',
                '\\' => '\\',
                '0' => '\0',
                _ => escaped
            };
        }
        else if (!IsAtEnd())
        {
            value = Advance();
        }
        else
        {
            AddErrorToken(start, "Unterminated character literal");
            return;
        }

        if (IsAtEnd() || Peek() != '\'')
        {
            AddErrorToken(start, "Unterminated character literal");
            return;
        }

        Advance(); // consume closing quote
        var lexeme = _source[start.._position];
        _tokens.Add(new Token(TokenType.CharLiteral, lexeme, value, start));
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

    private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_' || c == '$';

    private static bool IsIdentifierChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '$';

    private void AddToken(TokenType type, int start, string? lexeme = null)
    {
        lexeme ??= _source[start.._position];
        _tokens.Add(new Token(type, lexeme, null, start));
    }

    private void AddErrorToken(int start, string message)
    {
        _tokens.Add(new Token(TokenType.Error, message, null, start));
    }

    /// <summary>
    /// Captures a brace-balanced statement block from source starting at given position.
    /// Handles string literals, verbatim strings, char literals, and comments that may contain braces.
    /// </summary>
    /// <param name="source">The source string to capture from.</param>
    /// <param name="startPosition">The position to start capturing from.</param>
    /// <returns>A tuple of (block content including braces, end position), or null if not a valid block.</returns>
    public static (string Block, int EndPosition)? CaptureStatementBlock(string source, int startPosition)
    {
        var pos = startPosition;

        // Skip whitespace
        while (pos < source.Length && char.IsWhiteSpace(source[pos]))
        {
            pos++;
        }

        // Verify starts with {
        if (pos >= source.Length || source[pos] != '{')
        {
            return null;
        }

        var blockStart = pos;
        var depth = 0;

        while (pos < source.Length)
        {
            var c = source[pos];

            switch (c)
            {
                case '{':
                    depth++;
                    pos++;
                    break;

                case '}':
                    depth--;
                    pos++;
                    if (depth == 0)
                    {
                        // Successfully captured the block
                        return (source[blockStart..pos], pos);
                    }
                    break;

                case '"':
                    // Check for verbatim string (@"...")
                    if (pos > 0 && source[pos - 1] == '@')
                    {
                        pos++; // skip opening quote
                        // Verbatim string: ends at " not followed by another "
                        while (pos < source.Length)
                        {
                            if (source[pos] == '"')
                            {
                                pos++;
                                if (pos >= source.Length || source[pos] != '"')
                                {
                                    break; // End of verbatim string
                                }
                                pos++; // Skip escaped ""
                            }
                            else
                            {
                                pos++;
                            }
                        }
                    }
                    else
                    {
                        // Regular string literal
                        pos++; // skip opening quote
                        while (pos < source.Length && source[pos] != '"')
                        {
                            if (source[pos] == '\\' && pos + 1 < source.Length)
                            {
                                pos += 2; // Skip escape sequence
                            }
                            else
                            {
                                pos++;
                            }
                        }
                        if (pos < source.Length)
                        {
                            pos++; // skip closing quote
                        }
                    }
                    break;

                case '\'':
                    // Character literal
                    pos++; // skip opening quote
                    while (pos < source.Length && source[pos] != '\'')
                    {
                        if (source[pos] == '\\' && pos + 1 < source.Length)
                        {
                            pos += 2; // Skip escape sequence
                        }
                        else
                        {
                            pos++;
                        }
                    }
                    if (pos < source.Length)
                    {
                        pos++; // skip closing quote
                    }
                    break;

                case '/':
                    if (pos + 1 < source.Length)
                    {
                        if (source[pos + 1] == '/')
                        {
                            // Single-line comment
                            pos += 2;
                            while (pos < source.Length && source[pos] != '\n')
                            {
                                pos++;
                            }
                        }
                        else if (source[pos + 1] == '*')
                        {
                            // Multi-line comment
                            pos += 2;
                            while (pos + 1 < source.Length && !(source[pos] == '*' && source[pos + 1] == '/'))
                            {
                                pos++;
                            }
                            if (pos + 1 < source.Length)
                            {
                                pos += 2; // skip */
                            }
                        }
                        else
                        {
                            pos++;
                        }
                    }
                    else
                    {
                        pos++;
                    }
                    break;

                default:
                    pos++;
                    break;
            }
        }

        // Unbalanced braces
        return null;
    }
}
