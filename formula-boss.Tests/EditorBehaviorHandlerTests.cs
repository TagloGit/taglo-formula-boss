using FormulaBoss.UI;

using ICSharpCode.AvalonEdit;

using Xunit;

namespace FormulaBoss.Tests;

public class EditorBehaviorHandlerTests
{
    private static void RunOnSta(Action action)
    {
        Exception? caught = null;
        var thread = new Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex) { caught = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (caught != null)
        {
            throw caught;
        }
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
            var (editor, handler) = CreateEditor($"x{ch}y", 1);
            Assert.True(handler.TrySkipClosingChar(ch));
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Does_not_skip_when_char_differs() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x)y", 1);
            Assert.False(handler.TrySkipClosingChar(']'));
            Assert.Equal(1, editor.CaretOffset);
        });

        [Fact]
        public void Does_not_skip_opening_chars() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("x(y", 1);
            Assert.False(handler.TrySkipClosingChar('('));
        });

        [Fact]
        public void Does_not_skip_at_end_of_document() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x)", 2);
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
            var (editor, handler) = CreateEditor("x", 1);
            editor.Document.Insert(1, open.ToString());
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar(open);
            Assert.Contains(close.ToString(), editor.Text);
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Does_nothing_for_non_brace_char() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("abc", 3);
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
            var (editor, handler) = CreateEditor(text, text.IndexOf('{') + 1);
            handler.DeIndentOpenBrace();
            Assert.Equal("foo\r\n{}", editor.Text);
        });

        [Fact]
        public void Does_not_deindent_when_already_matching() => RunOnSta(() =>
        {
            var text = "  foo\r\n  {}";
            var (editor, handler) = CreateEditor(text, text.IndexOf('{') + 1);
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
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBraceBlock());
            Assert.Equal("  {\r\n    \r\n  }", editor.Text);
            Assert.Equal("  {\r\n    ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Returns_false_when_not_between_braces() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("abc", 1);
            Assert.False(handler.TryExpandBraceBlock());
        });

        [Fact]
        public void Returns_false_at_document_start() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("{}", 0);
            Assert.False(handler.TryExpandBraceBlock());
        });

        [Fact]
        public void Returns_false_at_document_end() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("{}", 2);
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
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBeforeClosingParen());
            Assert.Equal("foo(x\r\n  )", editor.Text);
        });

        [Fact]
        public void Expands_before_closing_bracket() => RunOnSta(() =>
        {
            var text = "arr[i]";
            var caretPos = text.IndexOf(']');
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);

            Assert.True(handler.TryExpandBeforeClosingParen());
            Assert.Equal("arr[i\r\n  ]", editor.Text);
        });

        [Fact]
        public void Returns_false_when_not_before_closer() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("abc", 1);
            Assert.False(handler.TryExpandBeforeClosingParen());
        });

        [Fact]
        public void Returns_false_at_document_end() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("foo()", 5);
            Assert.False(handler.TryExpandBeforeClosingParen());
        });
    }

    public class SmartHome : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Jumps_to_first_non_whitespace() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, text.Length);
            handler.SmartHome(false);
            Assert.Equal(4, editor.CaretOffset);
        });

        [Fact]
        public void Jumps_to_column_zero_when_at_first_non_whitespace() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, 4);
            handler.SmartHome(false);
            Assert.Equal(0, editor.CaretOffset);
        });

        [Fact]
        public void Toggles_back_to_first_non_whitespace() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, 0);
            handler.SmartHome(false);
            Assert.Equal(4, editor.CaretOffset);
        });

        [Fact]
        public void Extends_selection_with_shift() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, text.Length);
            handler.SmartHome(true);
            Assert.Equal(4, editor.CaretOffset);
            Assert.True(editor.SelectionLength > 0);
        });

        [Fact]
        public void Shift_home_twice_extends_selection_to_column_zero() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, text.Length);
            // First Shift+Home: select from end to first non-ws (col 4)
            handler.SmartHome(true);
            Assert.Equal(4, editor.CaretOffset);
            Assert.Equal(4, editor.SelectionStart);
            Assert.Equal(5, editor.SelectionLength);
            // Second Shift+Home: extend selection to col 0, anchor stays at end
            handler.SmartHome(true);
            Assert.Equal(0, editor.CaretOffset);
            Assert.Equal(0, editor.SelectionStart);
            Assert.Equal(9, editor.SelectionLength);
        });

        [Fact]
        public void Works_on_line_with_no_indent() => RunOnSta(() =>
        {
            var text = "hello";
            var (editor, handler) = CreateEditor(text, 3);
            handler.SmartHome(false);
            Assert.Equal(0, editor.CaretOffset);
        });
    }

    public class SmartBackspace : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Removes_to_previous_indent_stop() => RunOnSta(() =>
        {
            var text = "      x";
            // Caret at column 6 (6 spaces), indent size 2 -> should go to 4
            var (editor, handler) = CreateEditor(text, 6, indentSize: 2);
            Assert.True(handler.TrySmartBackspace());
            Assert.Equal("    x", editor.Text);
            Assert.Equal(4, editor.CaretOffset);
        });

        [Fact]
        public void Removes_to_zero_when_within_first_indent() => RunOnSta(() =>
        {
            var text = "  x";
            var (editor, handler) = CreateEditor(text, 2, indentSize: 4);
            Assert.True(handler.TrySmartBackspace());
            Assert.Equal("x", editor.Text);
            Assert.Equal(0, editor.CaretOffset);
        });

        [Fact]
        public void Does_nothing_when_not_in_leading_whitespace() => RunOnSta(() =>
        {
            var text = "  ab";
            var (editor, handler) = CreateEditor(text, 3, indentSize: 2);
            Assert.False(handler.TrySmartBackspace());
            Assert.Equal("  ab", editor.Text);
        });

        [Fact]
        public void Does_nothing_at_start_of_document() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("  x", 0, indentSize: 2);
            Assert.False(handler.TrySmartBackspace());
        });

        [Fact]
        public void Does_nothing_at_column_zero_of_line() => RunOnSta(() =>
        {
            var text = "foo\r\n  bar";
            // Caret at start of second line (column 0)
            var (_, handler) = CreateEditor(text, 5, indentSize: 2);
            Assert.False(handler.TrySmartBackspace());
        });
    }

    public class PairedDelete : EditorBehaviorHandlerTests
    {
        [Theory]
        [InlineData("()", 1)]
        [InlineData("[]", 1)]
        [InlineData("{}", 1)]
        [InlineData("\"\"", 1)]
        [InlineData("``", 1)]
        public void Deletes_both_chars_of_empty_pair(string pair, int caret) => RunOnSta(() =>
        {
            var text = "x" + pair + "y";
            var (editor, handler) = CreateEditor(text, caret + 1);
            Assert.True(handler.TryDeletePairedChars());
            Assert.Equal("xy", editor.Text);
        });

        [Fact]
        public void Does_not_delete_when_pair_is_not_empty() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("(a)", 2);
            Assert.False(handler.TryDeletePairedChars());
        });

        [Fact]
        public void Does_not_delete_at_start_of_document() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("()", 0);
            Assert.False(handler.TryDeletePairedChars());
        });

        [Fact]
        public void Does_not_delete_at_end_of_document() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("()", 2);
            Assert.False(handler.TryDeletePairedChars());
        });

        [Fact]
        public void Does_not_delete_mismatched_pair() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("(]", 1);
            Assert.False(handler.TryDeletePairedChars());
        });
    }

    public class SmartQuoteSuppression : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Auto_closes_quote_before_whitespace() => RunOnSta(() =>
        {
            // Simulate: user typed " and caret is after it, next char is space
            var (editor, handler) = CreateEditor("x y", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"\" y", editor.Text);
        });

        [Fact]
        public void Auto_closes_quote_at_end_of_document() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"\"", editor.Text);
        });

        [Fact]
        public void Suppresses_quote_before_letter() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("xy", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"y", editor.Text); // no closing quote added
        });

        [Fact]
        public void Suppresses_quote_before_digit() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x5", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"5", editor.Text);
        });

        [Fact]
        public void Suppresses_quote_before_underscore() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x_y", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"_y", editor.Text);
        });

        [Fact]
        public void Suppresses_quote_before_another_quote() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x\"", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"\"", editor.Text); // no extra quote
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Suppresses_when_inside_open_quote() => RunOnSta(() =>
        {
            // User has: "some text| and types " — should NOT auto-close
            var (editor, handler) = CreateEditor("\"some text ", 11);
            editor.Document.Insert(11, "\"");
            editor.CaretOffset = 12;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("\"some text \"", editor.Text); // no extra quote
        });

        [Fact]
        public void Auto_closes_when_previous_quotes_are_balanced() => RunOnSta(() =>
        {
            // User has: "a" | and types " — should auto-close (balanced)
            var (editor, handler) = CreateEditor("\"a\" ", 4);
            editor.Document.Insert(4, "\"");
            editor.CaretOffset = 5;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("\"a\" \"\"", editor.Text);
        });

        [Fact]
        public void Suppresses_backtick_when_inside_open_backtick() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("`expr ", 6);
            editor.Document.Insert(6, "`");
            editor.CaretOffset = 7;
            handler.AutoInsertClosingChar('`');
            Assert.Equal("`expr `", editor.Text);
        });

        [Fact]
        public void Auto_closes_quote_before_closing_bracket() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x)", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"\")", editor.Text);
        });

        [Fact]
        public void Auto_closes_quote_before_comma() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x,", 1);
            editor.Document.Insert(1, "\"");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('"');
            Assert.Equal("x\"\",", editor.Text);
        });

        [Fact]
        public void Auto_closes_backtick_before_whitespace() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("x y", 1);
            editor.Document.Insert(1, "`");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('`');
            Assert.Equal("x`` y", editor.Text);
        });

        [Fact]
        public void Suppresses_backtick_before_letter() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("xy", 1);
            editor.Document.Insert(1, "`");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('`');
            Assert.Equal("x`y", editor.Text);
        });

        [Fact]
        public void Does_not_suppress_paren_before_letter() => RunOnSta(() =>
        {
            // Smart suppression only applies to quotes, not brackets
            var (editor, handler) = CreateEditor("xy", 1);
            editor.Document.Insert(1, "(");
            editor.CaretOffset = 2;
            handler.AutoInsertClosingChar('(');
            Assert.Equal("x()y", editor.Text);
        });
    }

    public class TabIndent : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Inserts_spaces_to_next_tab_stop_no_selection() => RunOnSta(() =>
        {
            // Caret at column 1, indent size 4 → insert 3 spaces to reach column 4
            var (editor, handler) = CreateEditor("ab", 1, indentSize: 4);
            handler.Indent();
            Assert.Equal("a   b", editor.Text);
            Assert.Equal(4, editor.CaretOffset);
        });

        [Fact]
        public void Inserts_full_indent_at_tab_stop() => RunOnSta(() =>
        {
            // Caret at column 0, indent size 2 → insert 2 spaces
            var (editor, handler) = CreateEditor("ab", 0, indentSize: 2);
            handler.Indent();
            Assert.Equal("  ab", editor.Text);
            Assert.Equal(2, editor.CaretOffset);
        });

        [Fact]
        public void Indents_all_selected_lines_multiline() => RunOnSta(() =>
        {
            var text = "aaa\r\nbbb\r\nccc";
            var (editor, handler) = CreateEditor(text, 0, indentSize: 2);
            // Select lines 1 and 2 (first two lines)
            editor.Select(0, "aaa\r\nbbb".Length);
            handler.Indent();
            Assert.Equal("  aaa\r\n  bbb\r\nccc", editor.Text);
            Assert.True(editor.SelectionLength > 0);
        });

        [Fact]
        public void Single_line_selection_inserts_to_tab_stop() => RunOnSta(() =>
        {
            // Single-line selection behaves like no selection: inserts spaces at caret
            var (editor, handler) = CreateEditor("hello", 3, indentSize: 4);
            editor.Select(1, 2); // select "el", caret at 3
            handler.Indent();
            // Caret was at col 3, next tab stop at 4 → insert 1 space
            Assert.Equal("hel lo", editor.Text);
        });
    }

    public class ShiftTabDedent : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Removes_one_indent_level() => RunOnSta(() =>
        {
            var text = "    hello";
            var (editor, handler) = CreateEditor(text, 6, indentSize: 2);
            handler.Dedent();
            Assert.Equal("  hello", editor.Text);
        });

        [Fact]
        public void Removes_all_leading_whitespace_when_less_than_indent() => RunOnSta(() =>
        {
            var text = " hello";
            var (editor, handler) = CreateEditor(text, 3, indentSize: 4);
            handler.Dedent();
            Assert.Equal("hello", editor.Text);
        });

        [Fact]
        public void Dedents_all_selected_lines_multiline() => RunOnSta(() =>
        {
            var text = "  aaa\r\n  bbb\r\nccc";
            var (editor, handler) = CreateEditor(text, 0, indentSize: 2);
            editor.Select(0, "  aaa\r\n  bbb".Length);
            handler.Dedent();
            Assert.Equal("aaa\r\nbbb\r\nccc", editor.Text);
            Assert.True(editor.SelectionLength > 0);
        });

        [Fact]
        public void Does_nothing_on_line_with_no_indent() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("hello", 3, indentSize: 2);
            handler.Dedent();
            Assert.Equal("hello", editor.Text);
        });

        [Fact]
        public void Preserves_selection_after_dedent() => RunOnSta(() =>
        {
            var text = "  aaa\r\n  bbb\r\n  ccc";
            var (editor, handler) = CreateEditor(text, 0, indentSize: 2);
            editor.Select(0, "  aaa\r\n  bbb\r\n  ccc".Length);
            handler.Dedent();
            Assert.Equal("aaa\r\nbbb\r\nccc", editor.Text);
            // Selection should span all three lines
            Assert.Equal(0, editor.SelectionStart);
            Assert.Equal("aaa\r\nbbb\r\nccc".Length, editor.SelectionLength);
        });
    }

    public class StructuralEnterArrow : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Breaks_lambda_body_after_arrow_before_brace() => RunOnSta(() =>
        {
            // Caret before { (not between {}), so TryExpandAfterArrow fires
            var text = "x => {y}";
            var caretPos = text.IndexOf('{');
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Equal("x =>\r\n  {y}", editor.Text);
            Assert.Equal("x =>\r\n  ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Combined_arrow_brace_expansion() => RunOnSta(() =>
        {
            // Caret between {} with => before — combined expansion
            var text = "r => {}";
            var caretPos = text.IndexOf('}');
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Equal("r =>\r\n  {\r\n    \r\n  }", editor.Text);
            Assert.Equal("r =>\r\n  {\r\n    ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Combined_arrow_brace_preserves_trailing_content() => RunOnSta(() =>
        {
            var text = "r => {})`)";
            var caretPos = text.IndexOf('}');
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Equal("r =>\r\n  {\r\n    \r\n  })`)".Replace("    \r\n  }", "    \r\n  }"), editor.Text);
        });

        [Fact]
        public void Does_not_trigger_without_arrow() => RunOnSta(() =>
        {
            var text = "x = {}";
            var caretPos = text.IndexOf('}');
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            // Simple brace expansion (no arrow)
            Assert.Equal("x = {\r\n  \r\n}", editor.Text);
        });
    }

    public class StructuralEnterLetBinding : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Splits_let_bindings_onto_new_line() => RunOnSta(() =>
        {
            var text = "LET(x, 1, y, 2)";
            var caretPos = "LET(x, 1, ".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Equal("LET(x, 1,\r\n  y, 2)", editor.Text);
        });

        [Fact]
        public void Does_not_trigger_outside_let() => RunOnSta(() =>
        {
            var text = "FOO(x, 1, y)";
            var caretPos = "FOO(x, 1, ".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Contains("\r\n", editor.Text);
        });

        [Fact]
        public void Works_case_insensitive() => RunOnSta(() =>
        {
            var text = "let(a, 1, b, 2)";
            var caretPos = "let(a, 1, ".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.HandleEnter();
            Assert.Equal("let(a, 1,\r\n  b, 2)", editor.Text);
        });
    }

    public class AutoIndentEnter : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Inherits_indent_from_current_line() => RunOnSta(() =>
        {
            var text = "  hello world";
            var caretPos = "  hello".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.AutoIndentEnter();
            Assert.Equal("  hello\r\n  world", editor.Text);
            Assert.Equal("  hello\r\n  ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Indents_deeper_after_opener() => RunOnSta(() =>
        {
            var text = "  fn(";
            var (editor, handler) = CreateEditor(text, text.Length, indentSize: 2);
            handler.AutoIndentEnter();
            Assert.Equal("  fn(\r\n    ", editor.Text);
            Assert.Equal("  fn(\r\n    ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Splits_content_after_caret_to_new_line() => RunOnSta(() =>
        {
            var text = "fn(arg1, arg2)";
            var caretPos = "fn(arg1, ".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.AutoIndentEnter();
            Assert.Equal("fn(arg1,\r\narg2)", editor.Text);
        });

        [Fact]
        public void Caret_placed_before_moved_content() => RunOnSta(() =>
        {
            var text = "  foo bar";
            var caretPos = "  foo ".Length;
            var (editor, handler) = CreateEditor(text, caretPos, indentSize: 2);
            handler.AutoIndentEnter();
            Assert.Equal("  foo\r\n  bar", editor.Text);
            Assert.Equal("  foo\r\n  ".Length, editor.CaretOffset);
        });

        [Fact]
        public void Enter_at_end_of_line_creates_blank_indented_line() => RunOnSta(() =>
        {
            var text = "  hello";
            var (editor, handler) = CreateEditor(text, text.Length, indentSize: 2);
            handler.AutoIndentEnter();
            Assert.Equal("  hello\r\n  ", editor.Text);
        });
    }

    public class TrailingWhitespaceCleanup : EditorBehaviorHandlerTests
    {
        [Fact]
        public void Trims_blank_line_when_caret_leaves() => RunOnSta(() =>
        {
            var text = "hello\r\n    \r\nworld";
            var (editor, _) = CreateEditor(text, "hello\r\n  ".Length, indentSize: 2);
            // Move caret to line 3 to trigger cleanup of line 2
            editor.CaretOffset = "hello\r\n    \r\n".Length;
            Assert.Equal("hello\r\n\r\nworld", editor.Text);
        });

        [Fact]
        public void Does_not_trim_line_with_content() => RunOnSta(() =>
        {
            var text = "hello\r\n  x\r\nworld";
            var (editor, _) = CreateEditor(text, "hello\r\n  x".Length, indentSize: 2);
            editor.CaretOffset = "hello\r\n  x\r\n".Length;
            Assert.Equal("hello\r\n  x\r\nworld", editor.Text);
        });
    }

    public class SurroundSelection : EditorBehaviorHandlerTests
    {
        [Theory]
        [InlineData('(', ')')]
        [InlineData('[', ']')]
        [InlineData('{', '}')]
        [InlineData('"', '"')]
        [InlineData('`', '`')]
        public void Wraps_selection_with_pair(char open, char close) => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("xhelloy", 1);
            editor.Select(1, 5); // select "hello"
            Assert.True(handler.TrySurroundSelection(open));
            Assert.Equal($"x{open}hello{close}y", editor.Text);
            // Selection should be on "hello" inside the pair
            Assert.Equal(2, editor.SelectionStart);
            Assert.Equal(5, editor.SelectionLength);
        });

        [Fact]
        public void Does_nothing_without_selection() => RunOnSta(() =>
        {
            var (_, handler) = CreateEditor("hello", 2);
            Assert.False(handler.TrySurroundSelection('('));
        });

        [Fact]
        public void Does_nothing_for_non_opener() => RunOnSta(() =>
        {
            var (editor, handler) = CreateEditor("hello", 1);
            editor.Select(1, 3);
            Assert.False(handler.TrySurroundSelection('a'));
        });
    }
}
