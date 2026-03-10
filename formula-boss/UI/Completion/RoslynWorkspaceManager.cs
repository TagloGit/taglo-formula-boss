using System.Collections.Immutable;
using System.Diagnostics;

using FormulaBoss.Compilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Manages a persistent Roslyn <see cref="AdhocWorkspace" /> for intellisense completions.
///     Created when the floating editor opens; disposed when it closes.
/// </summary>
internal sealed class RoslynWorkspaceManager : IDisposable
{
    private readonly ProjectId _projectId;
    private readonly AdhocWorkspace _workspace;
    private DocumentId? _documentId;

    public RoslynWorkspaceManager()
    {
        var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
        _workspace = new AdhocWorkspace(host);

        var references = MetadataReferenceProvider.GetMetadataReferences();

        var projectInfo = ProjectInfo.Create(
            ProjectId.CreateNewId(),
            VersionStamp.Create(),
            "FormulaBoss.Intellisense",
            "FormulaBoss.Intellisense",
            LanguageNames.CSharp,
            metadataReferences: references,
            compilationOptions: new CSharpCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary));

        _workspace.AddProject(projectInfo);
        _projectId = projectInfo.Id;
    }

    public void Dispose() => _workspace.Dispose();

    private void UpdateDocument(string syntheticSource)
    {
        if (_documentId != null)
        {
            _workspace.TryApplyChanges(
                _workspace.CurrentSolution.RemoveDocument(_documentId));
        }

        _documentId = DocumentId.CreateNewId(_projectId);
        var documentInfo = DocumentInfo.Create(
            _documentId,
            "Intellisense.cs",
            sourceCodeKind: SourceCodeKind.Regular,
            loader: TextLoader.From(TextAndVersion.Create(
                SourceText.From(syntheticSource), VersionStamp.Create())));

        _workspace.TryApplyChanges(
            _workspace.CurrentSolution.AddDocument(documentInfo));
    }

    /// <summary>
    ///     Updates the synthetic document and returns Roslyn completions at the given caret offset.
    /// </summary>
    public async Task<IReadOnlyList<CompletionItem>> GetCompletionsAsync(
        string syntheticSource, int caretOffset, CancellationToken cancellationToken)
    {
        UpdateDocument(syntheticSource);

        var document = _workspace.CurrentSolution.GetDocument(_documentId);
        if (document == null)
        {
            return Array.Empty<CompletionItem>();
        }

        var completionService = CompletionService.GetService(document);
        if (completionService == null)
        {
            Debug.WriteLine("CompletionService not available");
            return Array.Empty<CompletionItem>();
        }

        try
        {
            var completions = await completionService.GetCompletionsAsync(
                document, caretOffset, cancellationToken: cancellationToken);

            return completions.ItemsList;
        }
        catch (OperationCanceledException)
        {
            return Array.Empty<CompletionItem>();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Roslyn completion error: {ex.Message}");
            return Array.Empty<CompletionItem>();
        }
    }

    /// <summary>
    ///     Updates the synthetic document and returns signature help info at the given caret offset
    ///     by walking the syntax tree and querying the semantic model for method symbols.
    /// </summary>
    public async Task<SignatureHelpModel?> GetSignatureHelpAsync(
        string syntheticSource, int caretOffset, CancellationToken cancellationToken)
    {
        UpdateDocument(syntheticSource);

        var document = _workspace.CurrentSolution.GetDocument(_documentId);
        if (document == null)
        {
            return null;
        }

        try
        {
            var root = await document.GetSyntaxRootAsync(cancellationToken);
            var semanticModel = await document.GetSemanticModelAsync(cancellationToken);
            if (root == null || semanticModel == null)
            {
                return null;
            }

            // Find the innermost argument list containing the caret.
            // The caret is placed right after the last character typed, so we look
            // at the token just before the caret to find the ( or , token that
            // belongs to the argument list.
            var token = root.FindToken(Math.Max(0, caretOffset - 1));
            var argumentList = token.Parent?.AncestorsAndSelf()
                .OfType<ArgumentListSyntax>()
                .FirstOrDefault();

            // Also try the token AT the caret (handles cases where caret is between args)
            if (argumentList == null)
            {
                token = root.FindToken(caretOffset);
                argumentList = token.Parent?.AncestorsAndSelf()
                    .OfType<ArgumentListSyntax>()
                    .FirstOrDefault();
            }

            if (argumentList?.Parent is not InvocationExpressionSyntax invocation)
            {
                return null;
            }

            // Determine active parameter index from comma count before caret
            var activeParameter = 0;
            foreach (var comma in argumentList.Arguments.GetSeparators())
            {
                if (comma.SpanStart < caretOffset)
                {
                    activeParameter++;
                }
            }

            // Get all candidate method symbols (overloads)
            var symbolInfo = semanticModel.GetSymbolInfo(invocation, cancellationToken);
            var candidates = new List<IMethodSymbol>();

            if (symbolInfo.Symbol is IMethodSymbol resolvedMethod)
            {
                candidates.Add(resolvedMethod);
            }

            foreach (var candidate in symbolInfo.CandidateSymbols.OfType<IMethodSymbol>())
            {
                if (!candidates.Contains(candidate, SymbolEqualityComparer.Default))
                {
                    candidates.Add(candidate);
                }
            }

            // Also check the member group for all overloads
            var memberGroup = semanticModel.GetMemberGroup(invocation.Expression, cancellationToken);
            foreach (var member in memberGroup.OfType<IMethodSymbol>())
            {
                if (!candidates.Contains(member, SymbolEqualityComparer.Default))
                {
                    candidates.Add(member);
                }
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            // Build overload models
            var overloads = candidates.Select(MapMethodToOverload).ToList();

            // Select the best overload (the resolved one, or the first that has enough parameters)
            var activeOverload = 0;
            if (symbolInfo.Symbol is IMethodSymbol)
            {
                activeOverload = 0; // resolved method is first
            }
            else
            {
                for (var i = 0; i < overloads.Count; i++)
                {
                    if (overloads[i].Parameters.Count > activeParameter)
                    {
                        activeOverload = i;
                        break;
                    }
                }
            }

            return new SignatureHelpModel(overloads, activeOverload, activeParameter);
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Roslyn signature help error: {ex.Message}");
            return null;
        }
    }

    private static SignatureOverload MapMethodToOverload(IMethodSymbol method)
    {
        var methodName = method.ContainingType != null
            ? $"{method.ContainingType.Name}.{method.Name}"
            : method.Name;

        // Filter out synthetic __ prefixed type names
        if (methodName.Contains("__"))
        {
            methodName = method.Name;
        }

        var xml = ResolveDocumentationXml(method);
        var summary = !string.IsNullOrEmpty(xml) ? ExtractXmlTag(xml, "summary") : null;

        var returnType = method.ReturnsVoid
            ? "void"
            : method.ReturnType.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);

        var parameters = method.Parameters.Select(p =>
        {
            var paramDoc = !string.IsNullOrEmpty(xml) ? ExtractParamDoc(xml, p.Name) : null;
            return new SignatureParameterInfo(
                p.Name,
                p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
                paramDoc);
        }).ToList();

        return new SignatureOverload(methodName, summary, parameters, returnType);
    }

    /// <summary>
    ///     Resolves the XML documentation for a method, walking up through overrides
    ///     and interface implementations to resolve <c>inheritdoc</c> tags.
    /// </summary>
    private static string? ResolveDocumentationXml(IMethodSymbol method)
    {
        var xml = method.GetDocumentationCommentXml();

        // If we have actual docs (not inheritdoc), return them
        if (!string.IsNullOrEmpty(xml) && !xml.Contains("<inheritdoc", StringComparison.OrdinalIgnoreCase))
        {
            return xml;
        }

        // Try the overridden method
        var overridden = method.OverriddenMethod;
        while (overridden != null)
        {
            xml = overridden.GetDocumentationCommentXml();
            if (!string.IsNullOrEmpty(xml) && !xml.Contains("<inheritdoc", StringComparison.OrdinalIgnoreCase))
            {
                return xml;
            }

            overridden = overridden.OverriddenMethod;
        }

        // Try interface implementations
        foreach (var iface in method.ContainingType?.AllInterfaces ?? ImmutableArray<INamedTypeSymbol>.Empty)
        {
            foreach (var ifaceMember in iface.GetMembers().OfType<IMethodSymbol>())
            {
                var impl = method.ContainingType?.FindImplementationForInterfaceMember(ifaceMember);
                if (SymbolEqualityComparer.Default.Equals(impl, method))
                {
                    xml = ifaceMember.GetDocumentationCommentXml();
                    if (!string.IsNullOrEmpty(xml) &&
                        !xml.Contains("<inheritdoc", StringComparison.OrdinalIgnoreCase))
                    {
                        return xml;
                    }
                }
            }
        }

        return null;
    }

    private static string? ExtractXmlTag(string xml, string tagName)
    {
        var startTag = $"<{tagName}>";
        var endTag = $"</{tagName}>";
        var start = xml.IndexOf(startTag, StringComparison.Ordinal);
        var end = xml.IndexOf(endTag, StringComparison.Ordinal);
        if (start < 0 || end < 0)
        {
            return null;
        }

        var content = xml[(start + startTag.Length)..end].Trim();
        // Strip inline XML tags like <see cref="..." />
        content = System.Text.RegularExpressions.Regex.Replace(content, @"<see\s+cref=""[^""]*""\s*/>", m =>
        {
            var cref = System.Text.RegularExpressions.Regex.Match(m.Value, @"cref=""([^""]*)""").Groups[1].Value;
            var lastDot = cref.LastIndexOf('.');
            return lastDot >= 0 ? cref[(lastDot + 1)..] : cref;
        });
        content = System.Text.RegularExpressions.Regex.Replace(content, @"<[^>]+>", "").Trim();
        return string.IsNullOrEmpty(content) ? null : content;
    }

    private static string? ExtractParamDoc(string xml, string paramName)
    {
        var pattern = $"<param name=\"{paramName}\">";
        var start = xml.IndexOf(pattern, StringComparison.Ordinal);
        if (start < 0)
        {
            return null;
        }

        var end = xml.IndexOf("</param>", start, StringComparison.Ordinal);
        if (end < 0)
        {
            return null;
        }

        var content = xml[(start + pattern.Length)..end].Trim();
        content = System.Text.RegularExpressions.Regex.Replace(content, @"<[^>]+>", "").Trim();
        return string.IsNullOrEmpty(content) ? null : content;
    }

    /// <summary>
    ///     Updates the synthetic document and returns Roslyn semantic diagnostics.
    /// </summary>
    public async Task<ImmutableArray<Diagnostic>> GetDiagnosticsAsync(
        string syntheticSource, CancellationToken cancellationToken)
    {
        UpdateDocument(syntheticSource);

        var document = _workspace.CurrentSolution.GetDocument(_documentId);
        if (document == null)
        {
            return ImmutableArray<Diagnostic>.Empty;
        }

        try
        {
            var semanticModel = await document.GetSemanticModelAsync(cancellationToken);
            if (semanticModel == null)
            {
                return ImmutableArray<Diagnostic>.Empty;
            }

            return semanticModel.GetDiagnostics(cancellationToken: cancellationToken)
                .Where(d => d.Severity is DiagnosticSeverity.Error or DiagnosticSeverity.Warning)
                .ToImmutableArray();
        }
        catch (OperationCanceledException)
        {
            return ImmutableArray<Diagnostic>.Empty;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Roslyn diagnostics error: {ex.Message}");
            return ImmutableArray<Diagnostic>.Empty;
        }
    }

    /// <summary>
    ///     Returns the display name of the type of the expression before the dot at the given
    ///     caret offset. Must be called after <see cref="GetCompletionsAsync" /> which loads
    ///     the document into the workspace.
    /// </summary>
    public async Task<string?> GetTypeBeforeDotAsync(int caretOffset, CancellationToken cancellationToken)
    {
        var document = _workspace.CurrentSolution.GetDocument(_documentId);
        if (document == null)
        {
            return null;
        }

        try
        {
            var semanticModel = await document.GetSemanticModelAsync(cancellationToken);
            var root = await document.GetSyntaxRootAsync(cancellationToken);
            if (semanticModel == null || root == null)
            {
                return null;
            }

            var token = root.FindToken(Math.Max(0, caretOffset - 1));

            var memberAccess = token.Parent?.AncestorsAndSelf()
                .OfType<MemberAccessExpressionSyntax>()
                .FirstOrDefault();

            if (memberAccess == null)
            {
                return null;
            }

            var typeInfo = semanticModel.GetTypeInfo(memberAccess.Expression, cancellationToken);
            return typeInfo.Type?.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Type lookup error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    ///     Pre-warms the workspace by requesting completions on a trivial document.
    /// </summary>
    public async Task WarmUpAsync()
    {
        try
        {
            await GetCompletionsAsync(
                "using System; class __W { void __M() { Console.} }",
                49,
                CancellationToken.None);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Roslyn warm-up error: {ex.Message}");
        }
    }
}
