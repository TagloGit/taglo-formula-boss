using System.Diagnostics;
using System.Reflection;

using ExcelDna.Integration;

using FormulaBoss.Compilation;
using FormulaBoss.Functions;
using FormulaBoss.Interception;

namespace FormulaBoss;

/// <summary>
/// Excel add-in entry point. Handles registration of static and dynamic UDFs.
/// </summary>
public sealed class AddIn : IExcelAddIn, IDisposable
{
    private DynamicCompiler? _compiler;
    private FormulaPipeline? _pipeline;
    private FormulaInterceptor? _interceptor;
    private bool _disposed;

    public void AutoOpen()
    {
        try
        {
            // Register static functions (ColorFunctions)
            RegisterStaticFunctions();

            // Defer event hookup until Excel is fully initialized
            // ExcelAsyncUtil.QueueAsMacro ensures we run after AutoOpen completes
            ExcelAsyncUtil.QueueAsMacro(InitializeInterception);

            Debug.WriteLine("Formula Boss add-in loaded successfully");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"AddIn.AutoOpen error: {ex}");
        }
    }

    private void InitializeInterception()
    {
        try
        {
            // Initialize the dynamic compilation infrastructure
            _compiler = new DynamicCompiler();
            _pipeline = new FormulaPipeline(_compiler);
            _interceptor = new FormulaInterceptor(_pipeline);

            // Start listening for worksheet changes
            _interceptor.Start();

            Debug.WriteLine("Formula Boss interception initialized");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"InitializeInterception error: {ex}");
        }
    }

    public void AutoClose()
    {
        Dispose();
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _interceptor?.Dispose();
        _interceptor = null;
        _pipeline = null;
        _compiler = null;
        _disposed = true;
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
