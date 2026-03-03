using System.Diagnostics;

using FormulaBoss.Compilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;

namespace FormulaBoss.UI.Completion;

/// <summary>
///     Manages a persistent Roslyn <see cref="AdhocWorkspace" /> for intellisense completions.
///     Created when the floating editor opens; disposed when it closes.
/// </summary>
internal sealed class RoslynWorkspaceManager : IDisposable
{
    private readonly AdhocWorkspace _workspace;
    private readonly ProjectId _projectId;
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
            compilationOptions: new Microsoft.CodeAnalysis.CSharp.CSharpCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary));

        _workspace.AddProject(projectInfo);
        _projectId = projectInfo.Id;
    }

    /// <summary>
    ///     Updates the synthetic document and returns Roslyn completions at the given caret offset.
    /// </summary>
    public async Task<IReadOnlyList<CompletionItem>> GetCompletionsAsync(
        string syntheticSource, int caretOffset, CancellationToken cancellationToken)
    {
        // Remove old document if present
        if (_documentId != null)
        {
            _workspace.TryApplyChanges(
                _workspace.CurrentSolution.RemoveDocument(_documentId));
        }

        // Add new document
        _documentId = DocumentId.CreateNewId(_projectId);
        var documentInfo = DocumentInfo.Create(
            _documentId,
            "Intellisense.cs",
            sourceCodeKind: SourceCodeKind.Regular,
            loader: TextLoader.From(TextAndVersion.Create(
                SourceText.From(syntheticSource), VersionStamp.Create())));

        _workspace.TryApplyChanges(
            _workspace.CurrentSolution.AddDocument(documentInfo));

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

            return completions?.ItemsList ?? (IReadOnlyList<CompletionItem>)Array.Empty<CompletionItem>();
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

    public void Dispose()
    {
        _workspace.Dispose();
    }
}
