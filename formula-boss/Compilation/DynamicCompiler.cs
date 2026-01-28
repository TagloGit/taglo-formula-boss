using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.Loader;

using ExcelDna.Integration;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FormulaBoss.Compilation;

/// <summary>
/// Compiles C# source code at runtime using Roslyn and registers the resulting UDFs with ExcelDNA.
/// </summary>
public class DynamicCompiler
{
    private static readonly string[] RequiredAssemblies =
    [
        "System.Runtime",
        "System.Private.CoreLib",
        "netstandard"
    ];

    /// <summary>
    /// Compiles and registers a test function to validate the Roslyn + ExcelDNA pipeline.
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

        RegisterFunctionsFromAssembly(assembly);
    }

    /// <summary>
    /// Compiles C# source code to an in-memory assembly.
    /// </summary>
    public static Assembly? CompileSource(string source)
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
                .Select(d => d.GetMessage(CultureInfo.InvariantCulture));

            Debug.WriteLine($"Compilation failed:\n{string.Join("\n", errors)}");
            return null;
        }

        ms.Seek(0, SeekOrigin.Begin);
        return AssemblyLoadContext.Default.LoadFromStream(ms);
    }

    /// <summary>
    /// Gets metadata references for common assemblies needed for compilation.
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
                    assemblyName.StartsWith("System.", StringComparison.Ordinal))
                {
                    references.Add(MetadataReference.CreateFromFile(assemblyPath));
                }
            }
        }

        return references;
    }

    /// <summary>
    /// Registers all public static methods from the compiled assembly as Excel UDFs.
    /// </summary>
    private static void RegisterFunctionsFromAssembly(Assembly assembly)
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
    /// Registers a single method as an Excel UDF.
    /// </summary>
    private static void RegisterMethod(MethodInfo method)
    {
        var funcAttr = new ExcelFunctionAttribute
        {
            Name = method.Name,
            Description = $"Dynamically compiled function: {method.Name}"
        };

        var parameters = method.GetParameters();
        var argAttrs = new List<object>();

        foreach (var param in parameters)
        {
            argAttrs.Add(new ExcelArgumentAttribute { Name = param.Name });
        }

        var delegateType = CreateDelegateType(method);
        var del = Delegate.CreateDelegate(delegateType, method);

        ExcelIntegration.RegisterDelegates(
            [del],
            [funcAttr],
            [argAttrs]);
    }

    /// <summary>
    /// Creates a Func delegate type matching the method signature.
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
