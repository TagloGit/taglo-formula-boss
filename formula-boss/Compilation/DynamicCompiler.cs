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

    // Trampoline dispatch: registered delegates call through this dictionary so that
    // re-editing a formula can swap the implementation without re-registering (which
    // would trigger an ExcelDNA "Repeated function name" warning popup).
    private readonly Dictionary<string, Delegate> _implementations = new();
    private readonly Dictionary<string, int> _registeredParamCounts = new();

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
    ///     Uses a trampoline pattern: the delegate registered with ExcelDNA dispatches
    ///     through <see cref="_implementations" />, so re-edits just swap the target
    ///     without calling <c>RegisterDelegates</c> again (avoiding the "Repeated
    ///     function name" warning popup).
    /// </summary>
    private void RegisterFunctionsFromAssembly(Assembly assembly, bool isMacroType)
    {
        foreach (var type in assembly.GetExportedTypes())
        {
            // DeclaredOnly so we don't pick up inherited object.Equals/ReferenceEquals as UDFs.
            var methods = type.GetMethods(
                BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);

            foreach (var method in methods)
            {
                try
                {
                    var paramCount = method.GetParameters().Length;
                    var implDelegate = Delegate.CreateDelegate(CreateDelegateType(method), method);

                    if (_registeredUdfs.Contains(method.Name)
                        && _registeredParamCounts.TryGetValue(method.Name, out var prevCount)
                        && prevCount == paramCount)
                    {
                        // Same param count — hot-swap the implementation; trampoline picks it up.
                        _implementations[method.Name] = implDelegate;
                        Debug.WriteLine($"Updated UDF implementation: {method.Name}");
                    }
                    else
                    {
                        // First registration (or param count changed) — register trampoline.
                        _implementations[method.Name] = implDelegate;
                        var trampoline = CreateTrampoline(method.Name, paramCount);
                        RegisterTrampoline(method.Name, trampoline, method, isMacroType);
                        _registeredUdfs.Add(method.Name);
                        _registeredParamCounts[method.Name] = paramCount;
                        Debug.WriteLine($"Registered dynamic UDF: {method.Name}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"RegisterUDF({method.Name})", ex);
                }
            }
        }
    }

    /// <summary>
    ///     Static version for test/spike functions — registers directly without trampolines.
    /// </summary>
    private static void RegisterFunctionsFromAssemblyStatic(Assembly assembly)
    {
        foreach (var type in assembly.GetExportedTypes())
        {
            var methods = type.GetMethods(
                BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);

            foreach (var method in methods)
            {
                try
                {
                    RegisterMethodDirect(method);
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
    ///     Registers a single method directly with ExcelDNA (no trampoline).
    ///     Used only for one-off test/spike functions that are never re-registered.
    /// </summary>
    private static void RegisterMethodDirect(MethodInfo method, bool isMacroType = false)
    {
        var funcAttr = new ExcelFunctionAttribute
        {
            Name = method.Name,
            Description = $"Dynamic UDF: {method.Name}",
            IsMacroType = isMacroType
        };

        var argAttrs = method.GetParameters()
            .Select(p => (object)new ExcelArgumentAttribute { Name = p.Name, AllowReference = true })
            .ToList();

        var delegateType = CreateDelegateType(method);
        var del = Delegate.CreateDelegate(delegateType, method);

        ExcelIntegration.RegisterDelegates(
            [del],
            [funcAttr],
            [argAttrs]);
    }

    /// <summary>
    ///     Creates a trampoline delegate that dispatches through <see cref="_implementations" />.
    ///     The trampoline is registered once with ExcelDNA; the underlying implementation
    ///     can be swapped without re-registering.
    /// </summary>
    private Delegate CreateTrampoline(string name, int paramCount)
    {
        // All generated UDF params/returns are object.
        return paramCount switch
        {
            0 => new Func<object>(
                () => ((Func<object>)_implementations[name])()),
            1 => new Func<object, object>(
                p0 => ((Func<object, object>)_implementations[name])(p0)),
            2 => new Func<object, object, object>(
                (p0, p1) => ((Func<object, object, object>)_implementations[name])(p0, p1)),
            3 => new Func<object, object, object, object>(
                (p0, p1, p2) => ((Func<object, object, object, object>)_implementations[name])(p0, p1, p2)),
            4 => new Func<object, object, object, object, object>(
                (p0, p1, p2, p3) =>
                    ((Func<object, object, object, object, object>)_implementations[name])(p0, p1, p2, p3)),
            _ => throw new NotSupportedException($"Methods with {paramCount} parameters not supported")
        };
    }

    /// <summary>
    ///     Registers a trampoline delegate with ExcelDNA, using parameter metadata from
    ///     the original compiled method.
    /// </summary>
    private static void RegisterTrampoline(
        string name, Delegate trampoline, MethodInfo originalMethod, bool isMacroType)
    {
        var funcAttr = new ExcelFunctionAttribute
        {
            Name = name,
            Description = $"Dynamic UDF: {name}",
            IsMacroType = isMacroType
        };

        var argAttrs = originalMethod.GetParameters()
            .Select(p => (object)new ExcelArgumentAttribute { Name = p.Name, AllowReference = true })
            .ToList();

        ExcelIntegration.RegisterDelegates(
            [trampoline],
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
