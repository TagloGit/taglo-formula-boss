namespace FormulaBoss.Parsing;

/// <summary>
/// Token types for the DSL lexer.
/// </summary>
public enum TokenType
{
    // Literals
    Identifier,
    Number,
    StringLiteral,

    // Operators
    Dot,
    Comma,
    Arrow,        // =>
    Plus,
    Minus,
    Star,
    Slash,
    Equal,        // ==
    NotEqual,     // !=
    Greater,
    Less,
    GreaterEqual, // >=
    LessEqual,    // <=
    And,          // &&
    Or,           // ||
    Not,          // !
    Colon,        // : (for range references like A1:B10)
    At,           // @ (for escape hatch in property access)

    // Delimiters
    LeftParen,
    RightParen,
    LeftBracket,
    RightBracket,

    // Special
    Eof,
    Error
}

/// <summary>
/// Represents a token produced by the lexer.
/// </summary>
/// <param name="Type">The token type.</param>
/// <param name="Lexeme">The source text that produced this token.</param>
/// <param name="Literal">For literals, the parsed value (double for numbers, string for strings).</param>
/// <param name="Position">Character position in the source where this token starts.</param>
public record Token(TokenType Type, string Lexeme, object? Literal, int Position);
