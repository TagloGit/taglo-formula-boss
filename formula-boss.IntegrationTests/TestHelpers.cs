using System.Reflection;
using System.Text;

using FormulaBoss.Interception;
using FormulaBoss.Parsing;
using FormulaBoss.Transpilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Helpers for compiling and executing generated UDF code in integration tests.
/// </summary>
public static class TestHelpers
{
    /// <summary>
    ///     Transpiles a DSL expression and compiles it to an assembly.
    ///     Returns the compiled _Core method that can be invoked directly with test data.
    /// </summary>
    /// <param name="dslExpression">The DSL expression (e.g., "data.cells.select(c => c.value).toArray()")</param>
    /// <returns>A TestCompilationResult containing the compiled method and metadata</returns>
    public static TestCompilationResult CompileExpression(string dslExpression)
    {
        // Parse
        var lexer = new Lexer(dslExpression);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);
        var expression = parser.Parse();

        if (expression == null || parser.Errors.Count > 0)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Parse error: {string.Join(", ", parser.Errors)}"
            };
        }

        // Transpile
        var transpiler = new CSharpTranspiler();
        var transpileResult = transpiler.Transpile(expression, dslExpression);

        // Compile
        var (assembly, errors) = CompileSource(transpileResult.SourceCode);

        if (assembly == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Compilation error:\n{string.Join("\n", errors)}",
                SourceCode = transpileResult.SourceCode
            };
        }

        // Find the _Core method
        var generatedType = assembly.GetType("GeneratedUdf");
        if (generatedType == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = "Could not find GeneratedUdf type in compiled assembly",
                SourceCode = transpileResult.SourceCode
            };
        }

        var coreMethodName = $"{transpileResult.MethodName}_Core";
        var coreMethod = generatedType.GetMethod(coreMethodName, BindingFlags.Public | BindingFlags.Static);

        if (coreMethod == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Could not find {coreMethodName} method in compiled assembly",
                SourceCode = transpileResult.SourceCode
            };
        }

        return new TestCompilationResult
        {
            Success = true,
            CoreMethod = coreMethod,
            MethodName = transpileResult.MethodName,
            RequiresObjectModel = transpileResult.RequiresObjectModel,
            SourceCode = transpileResult.SourceCode
        };
    }

    /// <summary>
    ///     Compiles C# source code to an in-memory assembly.
    /// </summary>
    private static (Assembly? Assembly, List<string> Errors) CompileSource(string sourceCode)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(sourceCode);

        var references = new List<MetadataReference>
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Enumerable).Assembly.Location),
            MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
            MetadataReference.CreateFromFile(Assembly.Load("System.Collections").Location),
            MetadataReference.CreateFromFile(Assembly.Load("System.Linq.Expressions").Location),
            MetadataReference.CreateFromFile(Assembly.Load("Microsoft.CSharp").Location),
            // Add reference to FormulaBoss for RuntimeHelpers
            MetadataReference.CreateFromFile(typeof(RuntimeHelpers).Assembly.Location)
        };

        // Add netstandard reference if available
        try
        {
            var netStandardPath = Path.Combine(
                Path.GetDirectoryName(typeof(object).Assembly.Location)!,
                "netstandard.dll");
            if (File.Exists(netStandardPath))
            {
                references.Add(MetadataReference.CreateFromFile(netStandardPath));
            }
        }
        catch
        {
            // Ignore if netstandard not found
        }

        var compilation = CSharpCompilation.Create(
            $"TestAssembly_{Guid.NewGuid():N}",
            [syntaxTree],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        using var ms = new MemoryStream();
        var result = compilation.Emit(ms);

        if (!result.Success)
        {
            var errors = result.Diagnostics
                .Where(d => d.Severity == DiagnosticSeverity.Error)
                .Select(d => d.ToString())
                .ToList();
            return (null, errors);
        }

        ms.Seek(0, SeekOrigin.Begin);
        var assembly = Assembly.Load(ms.ToArray());
        return (assembly, []);
    }

    /// <summary>
    ///     Executes a compiled _Core method with the given range (for object model path).
    /// </summary>
    public static object? ExecuteWithRange(MethodInfo coreMethod, dynamic range) => coreMethod.Invoke(null, [range]);

    /// <summary>
    ///     Executes a compiled _Core method with the given values array (for value-only path).
    /// </summary>
    public static object? ExecuteWithValues(MethodInfo coreMethod, object[,] values) =>
        coreMethod.Invoke(null, [values]);

    /// <summary>
    ///     Executes a compiled _Core method with the given values array and additional column name parameters.
    ///     Used for testing dynamic column name resolution.
    /// </summary>
    public static object? ExecuteWithValuesAndColumnNames(MethodInfo coreMethod, object[,] values, params string[] columnNames)
    {
        var args = new object[1 + columnNames.Length];
        args[0] = values;
        for (var i = 0; i < columnNames.Length; i++)
        {
            args[i + 1] = columnNames[i];
        }
        return coreMethod.Invoke(null, args);
    }

    /// <summary>
    ///     Transpiles a DSL expression with column bindings and compiles it to an assembly.
    ///     Returns the compiled _Core method that can be invoked directly with test data.
    /// </summary>
    /// <param name="dslExpression">The DSL expression</param>
    /// <param name="columnBindings">Column bindings from LET variables (e.g., "price" -> (tblSales, Price))</param>
    /// <returns>A TestCompilationResult containing the compiled method and metadata</returns>
    public static TestCompilationResult CompileExpressionWithColumnBindings(
        string dslExpression,
        Dictionary<string, Interception.ColumnBindingInfo> columnBindings)
    {
        // Parse
        var lexer = new Lexer(dslExpression);
        var tokens = lexer.ScanTokens();
        var parser = new Parser(tokens);
        var expression = parser.Parse();

        if (expression == null || parser.Errors.Count > 0)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Parse error: {string.Join(", ", parser.Errors)}"
            };
        }

        // Transpile with column bindings
        var transpiler = new CSharpTranspiler();
        var transpileResult = transpiler.Transpile(expression, dslExpression, null, columnBindings);

        // Compile
        var (assembly, errors) = CompileSource(transpileResult.SourceCode);

        if (assembly == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Compilation error:\n{string.Join("\n", errors)}",
                SourceCode = transpileResult.SourceCode
            };
        }

        // Find the _Core method
        var generatedType = assembly.GetType("GeneratedUdf");
        if (generatedType == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = "Could not find GeneratedUdf type in compiled assembly",
                SourceCode = transpileResult.SourceCode
            };
        }

        var coreMethodName = $"{transpileResult.MethodName}_Core";
        var coreMethod = generatedType.GetMethod(coreMethodName, BindingFlags.Public | BindingFlags.Static);

        if (coreMethod == null)
        {
            return new TestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Could not find {coreMethodName} method in compiled assembly",
                SourceCode = transpileResult.SourceCode
            };
        }

        return new TestCompilationResult
        {
            Success = true,
            CoreMethod = coreMethod,
            MethodName = transpileResult.MethodName,
            RequiresObjectModel = transpileResult.RequiresObjectModel,
            SourceCode = transpileResult.SourceCode,
            UsedColumnBindings = transpileResult.UsedColumnBindings
        };
    }
}

/// <summary>
///     Result of compiling a DSL expression for testing.
/// </summary>
public class TestCompilationResult
{
    public bool Success { get; init; }
    public string? ErrorMessage { get; init; }
    public MethodInfo? CoreMethod { get; init; }
    public string? MethodName { get; init; }
    public bool RequiresObjectModel { get; init; }
    public string? SourceCode { get; init; }
    public IReadOnlyList<string>? UsedColumnBindings { get; init; }

    /// <summary>
    ///     Helper to print diagnostic information when a test fails.
    /// </summary>
    public string GetDiagnostics()
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== Test Compilation Diagnostics ===");
        sb.AppendLine($"Success: {Success}");

        if (!string.IsNullOrEmpty(ErrorMessage))
        {
            sb.AppendLine($"Error: {ErrorMessage}");
        }

        if (!string.IsNullOrEmpty(SourceCode))
        {
            sb.AppendLine("\n=== Generated Source Code ===");
            sb.AppendLine(SourceCode);
        }

        return sb.ToString();
    }
}
