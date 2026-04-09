using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.Loader;

using ExcelDna.Integration;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

using Taglo.Excel.Common;

namespace FormulaBoss.Compilation;

/// <summary>
///     Compiles C# source code at runtime using Roslyn and registers the resulting UDFs with ExcelDNA.
/// </summary>
public class DynamicCompiler
{
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

        var references = MetadataReferenceProvider.GetMetadataReferences();

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
    ///     Gets the AssemblyLoadContext used by the host (ExcelDNA) so that generated code
    ///     can resolve types from host-loaded assemblies like FormulaBoss.Runtime.
    /// </summary>
    private static AssemblyLoadContext GetHostLoadContext()
    {
        return AssemblyLoadContext.GetLoadContext(typeof(ExcelFunctionAttribute).Assembly)
               ?? AssemblyLoadContext.Default;
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
                    Logger.Error($"RegisterUDF({method.Name})", ex);
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
                    Logger.Error($"RegisterUDF({method.Name})", ex);
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
            Name = method.Name,
            Description = $"Dynamic UDF: {method.Name}",
            IsMacroType = isMacroType
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
