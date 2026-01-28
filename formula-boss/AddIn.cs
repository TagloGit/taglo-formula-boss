using System.Diagnostics;
using System.Reflection;

using ExcelDna.Integration;

using FormulaBoss.Compilation;
using FormulaBoss.Functions;

namespace FormulaBoss;

/// <summary>
/// Excel add-in entry point. Handles registration of static and dynamic UDFs.
/// </summary>
public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            // Register static functions (ColorFunctions)
            RegisterStaticFunctions();

            // Phase 2: Compile and register a test dynamic UDF
            var stopwatch = Stopwatch.StartNew();
            DynamicCompiler.CompileAndRegisterTestFunction();
            stopwatch.Stop();

            Debug.WriteLine($"Dynamic compilation completed in {stopwatch.ElapsedMilliseconds}ms");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"AddIn.AutoOpen error: {ex}");
        }
    }

    public void AutoClose()
    {
        // Cleanup if needed
    }

    /// <summary>
    /// Manually register static UDF functions since we're using ExplicitRegistration.
    /// </summary>
    private static void RegisterStaticFunctions()
    {
        var methods = typeof(ColorFunctions).GetMethods(BindingFlags.Public | BindingFlags.Static);

        foreach (var method in methods)
        {
            var funcAttr = method.GetCustomAttribute<ExcelFunctionAttribute>();
            if (funcAttr == null)
            {
                continue;
            }

            var parameters = method.GetParameters();
            var argAttrs = new List<object>();

            foreach (var param in parameters)
            {
                var argAttr = param.GetCustomAttribute<ExcelArgumentAttribute>() ?? new ExcelArgumentAttribute();
                argAttrs.Add(argAttr);
            }

            // Create delegate from method
            var delegateType = CreateDelegateType(method);
            var del = Delegate.CreateDelegate(delegateType, method);

            ExcelIntegration.RegisterDelegates(
                [del],
                [funcAttr],
                [argAttrs]);
        }
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
