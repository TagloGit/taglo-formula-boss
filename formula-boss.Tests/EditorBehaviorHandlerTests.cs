using ICSharpCode.AvalonEdit;
using FormulaBoss.UI;
using Xunit;

namespace FormulaBoss.Tests;

public class EditorBehaviorHandlerTests
{
    private static void RunOnSta(Action action)
    {
        Exception? caught = null;
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { caught = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (caught != null) throw caught;
    }

    private static (TextEditor editor, EditorBehaviorHandler handler) CreateEditor(
        string text = "", int caretOffset = -1, int indentSize = 2)
    {
        var editor = new TextEditor();
        editor.Options.IndentationSize = indentSize;
        editor.Text = text;
        editor.CaretOffset = caretOffset >= 0 ? caretOffset : text.Length;
        var handler = new EditorBehaviorHandler(editor);
        return (editor, handler);
    }

    public class SkipClosingChar : EditorBehaviorHandlerTests
    {
        [Theory]
        [InlineData(')')]
        [InlineData(']')]
        [InlineData('}')]
        [InlineData('"')]
        [InlineData('`')]
        public void Skips_when_char_at_caret_matches(char ch) => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor($"x{ch}y", caretOffset: 1);
            Assert.True(handler.TrySkipClosingChar(ch));
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Does_not_skip_when_char_differs() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x)y", caretOffset: 1);
            Assert.False(handler.TrySkipClosingChar(']'));
            Assert.Equal(1, editor.CaretOffset);
        });

        [Fact]
        public void Does_not_skip_opening_chars() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("x(y", caretOffset: 1);
            Assert.False(handler.TrySkipClosingChar('('));
        });

        [Fact]
        public void Does_not_skip_at_end_of_document() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x)", caretOffset: 2);
            Assert.False(handler.TrySkipClosingChar(')'));
            Assert.Equal(2, editor.CaretOffset);
        });
    }

    public class AutoInsertClosingChar : EditorBehaviorHandlerTests
    {
        [Theory]
        [InlineData('(', ')')]
        [InlineData('[', ']')]
        [InlineData('{', '}')]
        [InlineData('"', '"')]
        [InlineData('`', '`')]
        public void Inserts_matching_closer(char open, char close) => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x", caretOffset: 1);
            editor.Document.Insert(1, open.ToString());
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar(open);
            Assert.Contains(close.ToString(), editor.Text);
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Does_nothing_for_non_brace_char() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("abc", caretOffset: 3);
            handler.AutoInsertClosingChar('a');
            Assert.Equal("abc", editor.Text);
        });
    }

    public class DeIndentOpenBrace : EditorBehaviorHandlerTests
    {
        [Fact]
        public void DeIndents_brace_to_match_previous_line() => RunOnSta(() =>
        {
            var text = "foo\r\n    {}";
            var (editor, handler) = CreateEditor(text, caretOffset: text.IndexOf('{') + 1);
            handler.DeIndentOpenBrace();
            Assert.Equal("foo\r\n{}", editor.Text);
        });

        [Fact]
        public void Does_not_deindent_when_already_matching() => RunOnSta(() =>
        {
            var text = "  foo\r\n  {}";
            var (editor, handler) = CreateEditor(text, caretOffset: text.IndexOf('{') + 1);
            handler.DeIndentOpenBrace();
            Assert.Equal("  foo\r\n  {}", editor.Text);
        });

        [Fact]
        public void Does_nothing_when_not_brace_line() => RunOnSta(() =>
        {
            var text = "foo\r\n    bar";
            var (editor, handler) = CreateEditor(text);
            handler.DeIndentOpenBrace();
            Assert.Equal(text, editor.Text);
        });
    }

    public class ExpandBraceBlock : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Expands_braces_on_enter() => RunOnSta(() =>
        {
            var text = "  {}";
            var caretPos = text.IndexOf('}');
            var (editor, handler) = CreateEditor(text, caretOffset: caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBraceBlock());
            Assert.Equal("  {\r\n    \r\n  }", editor.Text);
            Assert.Equal("  {\r\n    ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Returns_false_when_not_between_braces() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("abc", caretOffset: 1);
            Assert.False(handler.TryExpandBraceBlock());
        });

        [Fact]
        public void Returns_false_at_document_start() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("{}", caretOffset: 0);
            Assert.False(handler.TryExpandBraceBlock());
        });

        [Fact]
        public void Returns_false_at_document_end() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("{}", caretOffset: 2);
            Assert.False(handler.TryExpandBraceBlock());
        });
    }

    public class ExpandBeforeClosingParen : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Expands_before_closing_paren() => RunOnSta(() =>
        {
            var text = "foo(x)";
            var caretPos = text.IndexOf(')');
            var (editor, handler) = CreateEditor(text, caretOffset: caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBeforeClosingParen());
            Assert.Equal("foo(x\r\n  )", editor.Text);
        });

        [Fact]
        public void Expands_before_closing_bracket() => RunOnSta(() =>
        {
            var text = "arr[i]";
            var caretPos = text.IndexOf(']');
            var (editor, handler) = CreateEditor(text, caretOffset: caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBeforeClosingParen());
            Assert.Equal("arr[i\r\n  ]", editor.Text);
        });

        [Fact]
        public void Returns_false_when_not_before_closer() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("abc", caretOffset: 1);
            Assert.False(handler.TryExpandBeforeClosingParen());
        });

        [Fact]
        public void Returns_false_at_document_end() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("foo()", caretOffset: 5);
            Assert.False(handler.TryExpandBeforeClosingParen());
        });
    }
}
