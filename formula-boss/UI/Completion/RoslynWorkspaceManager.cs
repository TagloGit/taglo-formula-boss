using System.Collections.Immutable;
using System.Diagnostics;

using FormulaBoss.Compilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.CSharp;
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
