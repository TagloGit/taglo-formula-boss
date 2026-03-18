using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FormulaBoss.Compilation;

/// <summary>
///     Resolves public type names accessible from the default using directives
///     added to all generated UDF code. Used by <see cref="Transpilation.InputDetector" />
///     to avoid treating BCL type names (e.g. Regex, DateTime) as free variables.
/// </summary>
internal static class ImportedTypeNames
{
    /// <summary>
    ///     The namespaces added as using directives to all generated code.
    ///     Must match the usings emitted by <see cref="Transpilation.CodeEmitter.BuildSource" />.
    /// </summary>
    internal static readonly string[] ImportedNamespaces =
    {
        "System", "System.Collections", "System.Collections.Generic", "System.Linq", "System.Text",
        "System.Text.RegularExpressions", "FormulaBoss.Runtime"
    };

    private static readonly Lazy<HashSet<string>> Cached = new(Resolve);

    /// <summary>
    ///     Returns a cached set of all public type names accessible from the default usings.
    /// </summary>
    public static HashSet<string> Get() => Cached.Value;

    private static HashSet<string> Resolve()
    {
        var result = new HashSet<string>();

        try
        {
            var references = MetadataReferenceProvider.GetMetadataReferences();
            var compilation = CSharpCompilation.Create("TypeProbe", references: references);

            foreach (var ns in ImportedNamespaces)
            {
                var nsSymbol = FindNamespace(compilation.GlobalNamespace, ns);
                if (nsSymbol == null)
                {
                    continue;
                }

                foreach (var type in nsSymbol.GetTypeMembers())
                {
                    if (type.DeclaredAccessibility == Microsoft.CodeAnalysis.Accessibility.Public)
                    {
                        result.Add(type.Name);
                    }
                }
            }
        }
        catch
        {
            // Fall back gracefully — InputDetector still has C# keywords and Runtime types
        }

        return result;
    }

    private static INamespaceSymbol? FindNamespace(INamespaceSymbol root, string qualifiedName)
    {
        var current = root;
        foreach (var part in qualifiedName.Split('.'))
        {
            current = current.GetNamespaceMembers().FirstOrDefault(n => n.Name == part);
            if (current == null)
            {
                return null;
            }
        }

        return current;
    }
}
