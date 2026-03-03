using System.Diagnostics;
using System.Reflection;

using Microsoft.CodeAnalysis;

namespace FormulaBoss.Compilation;

/// <summary>
///     Provides metadata references for Roslyn compilation and completion workspaces.
///     Shared by <see cref="DynamicCompiler" /> and the Roslyn intellisense workspace.
/// </summary>
internal static class MetadataReferenceProvider
{
    private static readonly string[] RequiredAssemblies =
    [
        "System.Runtime",
        "System.Private.CoreLib",
        "netstandard"
    ];

    /// <summary>
    ///     Gets metadata references for common assemblies needed for compilation.
    /// </summary>
    public static List<MetadataReference> GetMetadataReferences()
    {
        var references = new List<MetadataReference>();

        // Add reference to the runtime assemblies
        if (AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") is string trustedAssemblies)
        {
            foreach (var assemblyPath in trustedAssemblies.Split(Path.PathSeparator))
            {
                var assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);
                if (RequiredAssemblies.Contains(assemblyName) ||
                    assemblyName.StartsWith("System.", StringComparison.Ordinal) ||
                    assemblyName.StartsWith("Microsoft.CSharp", StringComparison.Ordinal))
                {
                    references.Add(MetadataReference.CreateFromFile(assemblyPath));
                }
            }
        }

        AddExcelDnaReference(references);
        AddFormulaBossReference(references);
        AddRuntimeReference(references);

        return references;
    }

    private static void AddFormulaBossReference(List<MetadataReference> references)
    {
        var formulaBossAssembly = typeof(RuntimeHelpers).Assembly;

        if (!string.IsNullOrEmpty(formulaBossAssembly.Location))
        {
            references.Add(MetadataReference.CreateFromFile(formulaBossAssembly.Location));
            Debug.WriteLine($"Using FormulaBoss from Location: {formulaBossAssembly.Location}");
            return;
        }

        try
        {
            var assemblyBytes = GetAssemblyBytesFromMemory(formulaBossAssembly);
            if (assemblyBytes != null)
            {
                references.Add(MetadataReference.CreateFromImage(assemblyBytes));
                Debug.WriteLine("Created FormulaBoss reference from memory image");
                return;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to get FormulaBoss assembly bytes from memory: {ex.Message}");
        }

        Debug.WriteLine("WARNING: Could not add FormulaBoss assembly reference - compilation may fail");
    }

    private static void AddRuntimeReference(List<MetadataReference> references)
    {
        var runtimeAssembly = typeof(Runtime.ExcelValue).Assembly;

        if (!string.IsNullOrEmpty(runtimeAssembly.Location))
        {
            references.Add(MetadataReference.CreateFromFile(runtimeAssembly.Location));
            Debug.WriteLine($"Using FormulaBoss.Runtime from Location: {runtimeAssembly.Location}");
            return;
        }

        try
        {
            var assemblyBytes = GetAssemblyBytesFromMemory(runtimeAssembly);
            if (assemblyBytes != null)
            {
                references.Add(MetadataReference.CreateFromImage(assemblyBytes));
                Debug.WriteLine("Created FormulaBoss.Runtime reference from memory image");
                return;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to get Runtime assembly bytes from memory: {ex.Message}");
        }

        Debug.WriteLine("WARNING: Could not add FormulaBoss.Runtime assembly reference");
    }

    private static void AddExcelDnaReference(List<MetadataReference> references)
    {
        var excelDnaAssembly = typeof(ExcelDna.Integration.ExcelFunctionAttribute).Assembly;

        if (!string.IsNullOrEmpty(excelDnaAssembly.Location))
        {
            references.Add(MetadataReference.CreateFromFile(excelDnaAssembly.Location));
            Debug.WriteLine($"Using ExcelDNA from Location: {excelDnaAssembly.Location}");
            return;
        }

        var searchPaths = new List<string>
        {
            AppDomain.CurrentDomain.BaseDirectory,
            Path.GetDirectoryName(typeof(DynamicCompiler).Assembly.Location) ?? "",
            Environment.CurrentDirectory
        };

        var nugetCache = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".nuget", "packages", "exceldna.integration");

        if (Directory.Exists(nugetCache))
        {
            foreach (var versionDir in Directory.GetDirectories(nugetCache))
            {
                searchPaths.Add(Path.Combine(versionDir, "lib", "net6.0-windows7.0"));
                searchPaths.Add(Path.Combine(versionDir, "lib", "net6.0"));
                searchPaths.Add(Path.Combine(versionDir, "lib", "netstandard2.0"));
            }
        }

        foreach (var basePath in searchPaths)
        {
            if (string.IsNullOrEmpty(basePath))
            {
                continue;
            }

            var dllPath = Path.Combine(basePath, "ExcelDna.Integration.dll");
            if (File.Exists(dllPath))
            {
                references.Add(MetadataReference.CreateFromFile(dllPath));
                Debug.WriteLine($"Found ExcelDna.Integration.dll at: {dllPath}");
                return;
            }
        }

        try
        {
            var assemblyBytes = GetAssemblyBytesFromMemory(excelDnaAssembly);
            if (assemblyBytes != null)
            {
                references.Add(MetadataReference.CreateFromImage(assemblyBytes));
                Debug.WriteLine("Created ExcelDNA reference from memory image");
                return;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to get assembly bytes from memory: {ex.Message}");
        }

        Debug.WriteLine("WARNING: Could not add ExcelDNA assembly reference - compilation may fail");
    }

    internal static byte[]? GetAssemblyBytesFromMemory(Assembly assembly)
    {
        try
        {
            var module = assembly.ManifestModule;
            var fullyQualifiedName = module.FullyQualifiedName;

            if (File.Exists(fullyQualifiedName))
            {
                return File.ReadAllBytes(fullyQualifiedName);
            }

            var peImageField = typeof(Assembly).GetField("_peImage",
                BindingFlags.NonPublic | BindingFlags.Instance);

            if (peImageField != null && peImageField.GetValue(assembly) is byte[] peImage)
            {
                return peImage;
            }
        }
        catch
        {
            // Ignore reflection failures
        }

        return null;
    }
}
