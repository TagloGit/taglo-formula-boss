namespace FormulaBoss.UI.Completion;

/// <summary>
///     Orchestrates signature help requests: checks context, builds the synthetic document,
///     queries Roslyn, and returns a <see cref="SignatureHelpModel" />.
/// </summary>
internal sealed class SignatureHelpProvider
{
    private readonly RoslynWorkspaceManager _workspace;

    public SignatureHelpProvider(RoslynWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    /// <summary>
    ///     Returns signature help for the current cursor position, or null if not applicable.
    /// </summary>
    public async Task<SignatureHelpModel?> GetSignatureHelpAsync(
        string textUpToCaret, string fullText, WorkbookMetadata? metadata,
        CancellationToken cancellationToken)
    {
        // Only provide signature help inside DSL backtick regions
        if (!ContextResolver.IsInsideBackticks(textUpToCaret))
        {
            return null;
        }

        var (syntheticSource, caretOffset) = SyntheticDocumentBuilder.Build(
            fullText, textUpToCaret, metadata);

        return await _workspace.GetSignatureHelpAsync(
            syntheticSource, caretOffset, cancellationToken);
    }
}

/// <summary>Signature help data for display in the editor popup.</summary>
/// <param name="Overloads">All method overloads.</param>
/// <param name="ActiveOverloadIndex">Index of the best-matching overload.</param>
/// <param name="ActiveParameterIndex">Index of the parameter the cursor is on.</param>
internal sealed record SignatureHelpModel(
    IReadOnlyList<SignatureOverload> Overloads,
    int ActiveOverloadIndex,
    int ActiveParameterIndex);

/// <summary>A single method overload signature.</summary>
/// <param name="MethodName">Qualified method name (e.g. "RowCollection.Where").</param>
/// <param name="Summary">XML doc summary, or null.</param>
/// <param name="Parameters">Parameter list.</param>
/// <param name="ReturnType">Display string for the return type.</param>
internal sealed record SignatureOverload(
    string MethodName,
    string? Summary,
    IReadOnlyList<SignatureParameterInfo> Parameters,
    string ReturnType);

/// <summary>A single parameter in a method signature.</summary>
/// <param name="Name">Parameter name.</param>
/// <param name="Type">Display string for the parameter type.</param>
/// <param name="Description">XML doc description, or null.</param>
internal sealed record SignatureParameterInfo(
    string Name,
    string Type,
    string? Description);
