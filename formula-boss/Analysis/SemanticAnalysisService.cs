using System.Text;
using System.Text.RegularExpressions;

using FormulaBoss.Compilation;
using FormulaBoss.Transpilation;
using FormulaBoss.UI;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FormulaBoss.Analysis;

/// <summary>
///     Builds one-shot Roslyn compilations from user expressions + workbook metadata,
///     and exposes semantic model queries (type at offset, display formatting).
///     Used by the editor for hover type display.
/// </summary>
internal sealed class SemanticAnalysisService
{
    private static List<MetadataReference>? _cachedReferences;

    /// <summary>
    ///     Builds a <see cref="SemanticModel" /> for the given expression embedded in a typed
    ///     synthetic document. Returns null if the compilation cannot be created.
    /// </summary>
    /// <param name="expression">The user's DSL expression (without backticks).</param>
    /// <param name="isStatementBlock">True if the expression starts with '{'.</param>
    /// <param name="metadata">Workbook metadata with table/column info.</param>
    /// <param name="letBindings">LET variable bindings (name, value) from the enclosing formula.</param>
    /// <returns>
    ///     A <see cref="SemanticAnalysisResult" /> containing the semantic model and the offset
    ///     where the expression starts in the synthetic document, or null on failure.
    /// </returns>
    public SemanticAnalysisResult? BuildSemanticModel(
        string expression,
        bool isStatementBlock,
        WorkbookMetadata? metadata,
        IReadOnlyList<(string Name, string Value)>? letBindings)
    {
        var sb = new StringBuilder(1024);

        TypedDocumentBuilder.AppendUsings(sb);
        var tableTypeNames = TypedDocumentBuilder.EmitTableTypes(sb, metadata);

        sb.AppendLine("class __Ctx {");
        sb.AppendLine("void __M() {");

        // Declare LET binding variables with resolved types
        var declaredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (letBindings != null)
        {
            foreach (var (name, value) in letBindings)
            {
                var varName = name.Trim();
                if (!IsValidIdentifier(varName))
                {
                    continue;
                }

                var typeName = ResolveBindingType(value.Trim(), metadata, tableTypeNames);
                sb.AppendLine($"{typeName} {varName} = default!;");
                declaredNames.Add(varName);
            }
        }

        // Declare table and named range variables not already declared via LET
        if (metadata != null)
        {
            foreach (var tableName in metadata.TableNames)
            {
                if (!declaredNames.Contains(tableName) && IsValidIdentifier(tableName) &&
                    tableTypeNames.TryGetValue(tableName, out var tableTypeName))
                {
                    sb.AppendLine($"{tableTypeName} {tableName} = default!;");
                    declaredNames.Add(tableName);
                }
            }

            foreach (var name in metadata.NamedRanges)
            {
                if (!declaredNames.Contains(name) && IsValidIdentifier(name))
                {
                    sb.AppendLine($"ExcelArray {name} = default!;");
                    declaredNames.Add(name);
                }
            }
        }

        // Embed the expression
        int expressionStart;
        if (isStatementBlock)
        {
            sb.Append("object __userBlock() ");
            expressionStart = sb.Length;
            sb.AppendLine(expression);
            sb.AppendLine("var __result = __userBlock();");
        }
        else
        {
            sb.Append("var __result = ");
            expressionStart = sb.Length;
            sb.Append(expression);
            sb.AppendLine(";");
        }

        sb.AppendLine("}");
        sb.AppendLine("}");

        var source = sb.ToString();
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        _cachedReferences ??= MetadataReferenceProvider.GetMetadataReferences();

        var compilation = CSharpCompilation.Create(
            "SemanticAnalysis",
            new[] { syntaxTree },
            _cachedReferences,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        var semanticModel = compilation.GetSemanticModel(syntaxTree);
        return new SemanticAnalysisResult(semanticModel, expressionStart, expression.Length);
    }

    /// <summary>
    ///     Returns the type of the syntax node at the given offset within the expression.
    ///     The offset is relative to the start of the expression (not the synthetic document).
    /// </summary>
    public ITypeSymbol? GetTypeAtOffset(SemanticAnalysisResult result, int expressionOffset)
    {
        var absoluteOffset = result.ExpressionStart + expressionOffset;
        var root = result.SemanticModel.SyntaxTree.GetRoot();
        var token = root.FindToken(absoluteOffset);
        var node = token.Parent;

        if (node == null)
        {
            return null;
        }

        // Lambda parameter declaration: r in (r => ...)
        if (node is ParameterSyntax parameterSyntax)
        {
            var declaredSymbol = result.SemanticModel.GetDeclaredSymbol(parameterSyntax);
            return (declaredSymbol as IParameterSymbol)?.Type;
        }

        // Identifier name (variable reference, var keyword, etc.)
        if (node is IdentifierNameSyntax identifier)
        {
            // 'var' keyword in a declaration
            if (identifier.Identifier.Text == "var" &&
                node.Parent is VariableDeclarationSyntax varDecl)
            {
                var typeInfo = result.SemanticModel.GetTypeInfo(varDecl.Type);
                return typeInfo.Type;
            }

            // Regular identifiers — try symbol info first (gives declared type for variables)
            var symbolInfo = result.SemanticModel.GetSymbolInfo(identifier);
            if (symbolInfo.Symbol is ILocalSymbol local)
            {
                return local.Type;
            }

            if (symbolInfo.Symbol is IParameterSymbol param)
            {
                return param.Type;
            }

            // Fall back to type info
            var identTypeInfo = result.SemanticModel.GetTypeInfo(identifier);
            return identTypeInfo.Type;
        }

        // For other nodes, try type info
        var generalTypeInfo = result.SemanticModel.GetTypeInfo(node);
        return generalTypeInfo.Type;
    }

    /// <summary>
    ///     Formats an <see cref="ITypeSymbol" /> for hover display, stripping synthetic prefixes
    ///     and mapping to user-friendly names per the spec's formatting table.
    /// </summary>
    public string FormatTypeForDisplay(ITypeSymbol type, WorkbookMetadata? metadata)
    {
        var name = type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);

        // Handle nullable wrapper: strip trailing '?' for the inner check, re-add if needed
        var isNullable = name.EndsWith("?");
        var coreName = isNullable ? name[..^1] : name;

        // Synthetic table type: __<name>Table → ExcelTable (<originalName>)
        var tableMatch = Regex.Match(coreName, @"^__(.+)Table$");
        if (tableMatch.Success)
        {
            var safeName = tableMatch.Groups[1].Value;
            var originalName = ResolveOriginalTableName(safeName, metadata);
            return Wrap($"ExcelTable ({originalName})", isNullable);
        }

        // Synthetic row collection type: __<name>RowCollection → RowCollection (<originalName>)
        var rowCollMatch = Regex.Match(coreName, @"^__(.+)RowCollection$");
        if (rowCollMatch.Success)
        {
            var safeName = rowCollMatch.Groups[1].Value;
            var originalName = ResolveOriginalTableName(safeName, metadata);
            return Wrap($"RowCollection ({originalName})", isNullable);
        }

        // Synthetic row type: __<name>Row → Row {Col1, Col2, ...}
        var rowMatch = Regex.Match(coreName, @"^__(.+)Row$");
        if (rowMatch.Success)
        {
            var safeName = rowMatch.Groups[1].Value;
            var columns = ResolveColumnsForRow(safeName, metadata);
            if (columns != null)
            {
                var colDisplay = FormatColumnList(columns);
                return Wrap($"Row {{{colDisplay}}}", isNullable);
            }

            return Wrap("Row", isNullable);
        }

        // Strip generic type arguments that contain synthetic names
        if (coreName.Contains("__"))
        {
            coreName = RewriteGenericSyntheticTypes(coreName, metadata);
            return Wrap(coreName, isNullable);
        }

        return name;
    }

    private static string Wrap(string display, bool nullable) =>
        nullable ? display + "?" : display;

    /// <summary>
    ///     Resolves a sanitised table name back to the original table name from metadata.
    /// </summary>
    private static string ResolveOriginalTableName(string sanitisedName, WorkbookMetadata? metadata)
    {
        if (metadata == null)
        {
            return sanitisedName;
        }

        foreach (var (tableName, _) in metadata.TableColumns)
        {
            if (ColumnMapper.Sanitise(tableName) == sanitisedName)
            {
                return tableName;
            }
        }

        return sanitisedName;
    }

    /// <summary>
    ///     Resolves column names for a sanitised table name.
    /// </summary>
    private static IReadOnlyList<string>? ResolveColumnsForRow(
        string sanitisedName, WorkbookMetadata? metadata)
    {
        if (metadata == null)
        {
            return null;
        }

        foreach (var (tableName, columns) in metadata.TableColumns)
        {
            if (ColumnMapper.Sanitise(tableName) == sanitisedName)
            {
                return columns;
            }
        }

        return null;
    }

    private static string FormatColumnList(IReadOnlyList<string> columns)
    {
        const int maxColumns = 4;
        if (columns.Count <= maxColumns)
        {
            return string.Join(", ", columns);
        }

        return string.Join(", ", columns.Take(maxColumns)) + ", ...";
    }

    /// <summary>
    ///     Rewrites generic type strings that contain synthetic __ prefixed names.
    ///     E.g. "IEnumerable&lt;__SalesRow&gt;" → "IEnumerable&lt;Row {Date, Amount, ...}&gt;"
    /// </summary>
    private static string RewriteGenericSyntheticTypes(string typeName, WorkbookMetadata? metadata)
    {
        return Regex.Replace(typeName, @"__(\w+?)Row(?=[\s,>?\)]|$)", match =>
        {
            var safeName = match.Groups[1].Value;

            // Check if it's actually a RowCollection
            if (safeName.EndsWith("RowCollection"))
            {
                return match.Value;
            }

            var columns = ResolveColumnsForRow(safeName, metadata);
            if (columns != null)
            {
                return $"Row {{{FormatColumnList(columns)}}}";
            }

            return "Row";
        });
    }

    private static string ResolveBindingType(
        string value, WorkbookMetadata? metadata,
        Dictionary<string, string> tableTypeNames)
    {
        if (metadata == null)
        {
            return "ExcelValue";
        }

        if (tableTypeNames.TryGetValue(value, out var tableTypeName))
        {
            return tableTypeName;
        }

        foreach (var name in metadata.NamedRanges)
        {
            if (name.Equals(value, StringComparison.OrdinalIgnoreCase))
            {
                return "ExcelArray";
            }
        }

        return "ExcelValue";
    }

    private static bool IsValidIdentifier(string name) =>
        name.Length > 0 && (char.IsLetter(name[0]) || name[0] == '_') &&
        name.All(c => char.IsLetterOrDigit(c) || c == '_');
}

/// <summary>
///     Result of building a semantic model, containing the model and position mapping.
/// </summary>
internal sealed record SemanticAnalysisResult(
    SemanticModel SemanticModel,
    int ExpressionStart,
    int ExpressionLength);
