using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.Loader;

using ExcelDna.Integration;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FormulaBoss.Compilation;

/// <summary>
///     Compiles C# source code at runtime using Roslyn and registers the resulting UDFs with ExcelDNA.
/// </summary>
public class DynamicCompiler
{
    private static readonly string[] RequiredAssemblies =
    [
        "System.Runtime",
        "System.Private.CoreLib",
        "netstandard"
    ];

    private readonly HashSet<string> _registeredUdfs = [];

    /// <summary>
    ///     Compiles C# source code and registers all UDFs in it.
    ///     Returns a list of compilation errors, or empty list on success.
    /// </summary>
    public virtual List<string> CompileAndRegister(string source, bool isMacroType = false)
    {
        var (assembly, errors) = CompileSourceWithErrors(source);

        if (assembly == null)
        {
            return errors;
        }

        RegisterFunctionsFromAssembly(assembly, isMacroType);
        return [];
    }

    /// <summary>
    ///     Compiles and registers a spike UDF that references FormulaBoss.Runtime types directly.
    ///     Tests whether the ALC-based loading allows generated code to resolve Runtime types.
    ///     Call =ALC_SPIKE(A1) from Excel — should return the value, not #VALUE!.
    /// </summary>
    public static void CompileAndRegisterAlcSpike()
    {
        const string source = """
                              using FormulaBoss.Runtime;

                              public static class AlcSpikeTest
                              {
                                  public static object ALC_SPIKE(object raw)
                                  {
                                      // Extract value from ExcelReference if needed
                                      var values = FormulaBoss.RuntimeHelpers.GetValuesFromReference(raw);
                                      var wrapped = ExcelValue.Wrap(values);
                                      return wrapped.ToResult();
                                  }
                              }
                              """;

        var (assembly, errors) = CompileSourceWithErrors(source);
        if (assembly == null)
        {
            Debug.WriteLine($"ALC spike compilation failed:\n{string.Join("\n", errors)}");
            return;
        }

        RegisterFunctionsFromAssemblyStatic(assembly);
        Debug.WriteLine("ALC spike registered: =ALC_SPIKE(A1)");
    }

    /// <summary>
    ///     Compiles and registers a test function to validate the Roslyn + ExcelDNA pipeline.
    /// </summary>
    public static void CompileAndRegisterTestFunction()
    {
        const string source = """
                              using System;

                              namespace FormulaBoss.Generated
                              {
                                  public static class DynamicFunctions
                                  {
                                      public static double AddNumbers(double a, double b)
                                      {
                                          return a + b;
                                      }

                                      public static double MultiplyNumbers(double a, double b)
                                      {
                                          return a * b;
                                      }
                                  }
                              }
                              """;

        var assembly = CompileSource(source);
        if (assembly == null)
        {
            return;
        }

        RegisterFunctionsFromAssemblyStatic(assembly);
    }

    /// <summary>
    ///     Compiles C# source code to an in-memory assembly.
    /// </summary>
    public static Assembly? CompileSource(string source)
    {
        var (assembly, _) = CompileSourceWithErrors(source);
        return assembly;
    }

    /// <summary>
    ///     Compiles C# source code to an in-memory assembly, returning errors if compilation fails.
    /// </summary>
    private static (Assembly? Assembly, List<string> Errors) CompileSourceWithErrors(string source)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        var references = GetMetadataReferences();

        var compilation = CSharpCompilation.Create(
            $"FormulaBoss.Dynamic.{Guid.NewGuid():N}",
            [syntaxTree],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                .WithOptimizationLevel(OptimizationLevel.Release));

        using var ms = new MemoryStream();
        var result = compilation.Emit(ms);

        if (!result.Success)
        {
            var errors = result.Diagnostics
                .Where(d => d.Severity == DiagnosticSeverity.Error)
                .Select(d => d.GetMessage(CultureInfo.InvariantCulture))
                .ToList();

            Debug.WriteLine($"Compilation failed:\n{string.Join("\n", errors)}");
            return (null, errors);
        }

        ms.Seek(0, SeekOrigin.Begin);
        var alc = GetHostLoadContext();
        var assembly = alc.LoadFromStream(ms);
        return (assembly, []);
    }

    /// <summary>
    ///     Gets metadata references for common assemblies needed for compilation.
    /// </summary>
    private static List<MetadataReference> GetMetadataReferences()
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

        // Add ExcelDNA reference - handle embedded assembly case
        AddExcelDnaReference(references);

        // Add reference to formula-boss.dll for RuntimeHelpers
        AddFormulaBossReference(references);

        // Add reference to FormulaBoss.Runtime for wrapper types (ExcelValue, etc.)
        AddRuntimeReference(references);

        return references;
    }

    /// <summary>
    ///     Adds a reference to the formula-boss assembly containing RuntimeHelpers.
    /// </summary>
    private static void AddFormulaBossReference(List<MetadataReference> references)
    {
        var formulaBossAssembly = typeof(RuntimeHelpers).Assembly;

        // Try using Location first
        if (!string.IsNullOrEmpty(formulaBossAssembly.Location))
        {
            references.Add(MetadataReference.CreateFromFile(formulaBossAssembly.Location));
            Debug.WriteLine($"Using FormulaBoss from Location: {formulaBossAssembly.Location}");
            return;
        }

        // If Location is empty (packed assembly), try to read from memory
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

    /// <summary>
    ///     Gets the AssemblyLoadContext used by the host (ExcelDNA) so that generated code
    ///     can resolve types from host-loaded assemblies like FormulaBoss.Runtime.
    /// </summary>
    private static AssemblyLoadContext GetHostLoadContext()
    {
        return AssemblyLoadContext.GetLoadContext(typeof(ExcelFunctionAttribute).Assembly)
               ?? AssemblyLoadContext.Default;
    }

    /// <summary>
    ///     Adds a reference to the FormulaBoss.Runtime assembly for wrapper types.
    /// </summary>
    private static void AddRuntimeReference(List<MetadataReference> references)
    {
        var runtimeAssembly = typeof(FormulaBoss.Runtime.ExcelValue).Assembly;

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

    /// <summary>
    ///     Adds ExcelDNA assembly reference, handling the case where it's loaded from embedded resources.
    /// </summary>
    private static void AddExcelDnaReference(List<MetadataReference> references)
    {
        var excelDnaAssembly = typeof(ExcelFunctionAttribute).Assembly;

        // Try using Location first (works when not packed)
        if (!string.IsNullOrEmpty(excelDnaAssembly.Location))
        {
            references.Add(MetadataReference.CreateFromFile(excelDnaAssembly.Location));
            Debug.WriteLine($"Using ExcelDNA from Location: {excelDnaAssembly.Location}");
            return;
        }

        // When packed into XLL, Location is empty - search for the DLL
        // Build list of paths to search
        var searchPaths = new List<string>
        {
            AppDomain.CurrentDomain.BaseDirectory,
            Path.GetDirectoryName(typeof(DynamicCompiler).Assembly.Location) ?? "",
            Environment.CurrentDirectory
        };

        // Add NuGet packages cache paths
        var nugetCache = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".nuget", "packages", "exceldna.integration");

        if (Directory.Exists(nugetCache))
        {
            // Find available versions and search each
            foreach (var versionDir in Directory.GetDirectories(nugetCache))
            {
                // Try common target framework monikers
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

        // Last resort: read assembly bytes from memory
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

    /// <summary>
    ///     Attempts to read assembly bytes from a loaded assembly using reflection.
    /// </summary>
    private static byte[]? GetAssemblyBytesFromMemory(Assembly assembly)
    {
        // Try to get the raw assembly image using Module.ResolveSignature trick
        // or by reading from the ManifestModule
        try
        {
            var module = assembly.ManifestModule;

            // Use reflection to access internal/private methods that can give us the image
            var fullyQualifiedName = module.FullyQualifiedName;

            // If it's a file path, try reading it
            if (File.Exists(fullyQualifiedName))
            {
                return File.ReadAllBytes(fullyQualifiedName);
            }

            // Check if the assembly was loaded from a byte array (has no file backing)
            // In this case, we need to use Marshal to copy from the loaded image
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

    /// <summary>
    ///     Registers all public static methods from the compiled assembly as Excel UDFs.
    /// </summary>
    private void RegisterFunctionsFromAssembly(Assembly assembly, bool isMacroType)
    {
        foreach (var type in assembly.GetExportedTypes())
        {
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static);

            foreach (var method in methods)
            {
                // Skip if already registered
                if (_registeredUdfs.Contains(method.Name))
                {
                    Debug.WriteLine($"UDF already registered: {method.Name}");
                    continue;
                }

                try
                {
                    RegisterMethod(method, isMacroType);
                    _registeredUdfs.Add(method.Name);
                    Debug.WriteLine($"Registered dynamic UDF: {method.Name}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to register {method.Name}: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    ///     Static version for test functions (doesn't track UDFs).
    /// </summary>
    private static void RegisterFunctionsFromAssemblyStatic(Assembly assembly)
    {
        foreach (var type in assembly.GetExportedTypes())
        {
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static);

            foreach (var method in methods)
            {
                try
                {
                    RegisterMethod(method);
                    Debug.WriteLine($"Registered dynamic UDF: {method.Name}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to register {method.Name}: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    ///     Registers a single method as an Excel UDF.
    /// </summary>
    private static void RegisterMethod(MethodInfo method, bool isMacroType = false)
    {
        // Create function attribute - all dynamically compiled UDFs use the method name
        // IsMacroType = true is required for object model UDFs so that xlfReftext works
        var funcAttr = new ExcelFunctionAttribute
        {
            Name = method.Name, Description = $"Dynamic UDF: {method.Name}", IsMacroType = isMacroType
        };

        var parameters = method.GetParameters();
        var argAttrs = new List<object>();

        foreach (var param in parameters)
        {
            // All transpiled UDFs expect range references, so set AllowReference = true
            argAttrs.Add(new ExcelArgumentAttribute { Name = param.Name, AllowReference = true });
        }

        var delegateType = CreateDelegateType(method);
        var del = Delegate.CreateDelegate(delegateType, method);

        ExcelIntegration.RegisterDelegates(
            [del],
            [funcAttr],
            [argAttrs]);
    }

    /// <summary>
    ///     Creates a Func delegate type matching the method signature.
    /// </summary>
    private static Type CreateDelegateType(MethodInfo method)
    {
        var parameters = method.GetParameters();
        var paramTypes = parameters.Select(p => p.ParameterType).ToList();
        paramTypes.Add(method.ReturnType);

        return paramTypes.Count switch
        {
            1 => typeof(Func<>).MakeGenericType(paramTypes.ToArray()),
            2 => typeof(Func<,>).MakeGenericType(paramTypes.ToArray()),
            3 => typeof(Func<,,>).MakeGenericType(paramTypes.ToArray()),
            4 => typeof(Func<,,,>).MakeGenericType(paramTypes.ToArray()),
            5 => typeof(Func<,,,,>).MakeGenericType(paramTypes.ToArray()),
            _ => throw new NotSupportedException($"Methods with {parameters.Length} parameters not supported")
        };
    }
}
