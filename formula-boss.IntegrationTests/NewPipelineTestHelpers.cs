using System.Collections;
using System.Reflection;
using System.Text;

using FormulaBoss.Runtime;
using FormulaBoss.Transpilation;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FormulaBoss.IntegrationTests;

/// <summary>
///     Test helpers for the new InputDetector → CodeEmitter pipeline.
///     Generated code references FormulaBoss.Runtime types directly.
/// </summary>
public static class NewPipelineTestHelpers
{
    /// <summary>
    ///     Detects inputs and emits code for an expression, then compiles to an in-memory assembly.
    /// </summary>
    public static NewTestCompilationResult CompileExpression(string expression)
    {
        // Detect
        var detector = new InputDetector();
        DetectionResult detection;
        try
        {
            detection = detector.Detect(expression);
        }
        catch (Exception ex)
        {
            return new NewTestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Detection error: {ex.Message}"
            };
        }

        // Emit
        var emitter = new CodeEmitter();
        var transpileResult = emitter.Emit(detection, expression);

        // Compile
        var (assembly, errors) = CompileSource(transpileResult.SourceCode);

        if (assembly == null)
        {
            return new NewTestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Compilation error:\n{string.Join("\n", errors)}",
                SourceCode = transpileResult.SourceCode
            };
        }

        // Find the class and method
        var className = $"{transpileResult.MethodName}_Class";
        var generatedType = assembly.GetTypes().FirstOrDefault(t => t.Name == className);
        if (generatedType == null)
        {
            return new NewTestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Could not find {className} type in compiled assembly.\n" +
                               $"Available types: {string.Join(", ", assembly.GetTypes().Select(t => t.Name))}",
                SourceCode = transpileResult.SourceCode
            };
        }

        var method = generatedType.GetMethod(transpileResult.MethodName, BindingFlags.Public | BindingFlags.Static);
        if (method == null)
        {
            return new NewTestCompilationResult
            {
                Success = false,
                ErrorMessage = $"Could not find {transpileResult.MethodName} method in {className}",
                SourceCode = transpileResult.SourceCode
            };
        }

        return new NewTestCompilationResult
        {
            Success = true,
            Method = method,
            MethodName = transpileResult.MethodName,
            RequiresObjectModel = transpileResult.RequiresObjectModel,
            SourceCode = transpileResult.SourceCode,
            Detection = detection
        };
    }

    /// <summary>
    ///     Executes a compiled UDF method with a values array as input.
    ///     Sets up RuntimeHelpers.ToResultDelegate before invoking.
    /// </summary>
    public static object? ExecuteWithValues(MethodInfo method, object[,] values)
    {
        SetupDelegates();
        return method.Invoke(null, [values]);
    }

    /// <summary>
    ///     Executes a compiled UDF method with multiple inputs.
    /// </summary>
    public static object? ExecuteWithMultipleInputs(MethodInfo method, params object[] inputs)
    {
        SetupDelegates();
        return method.Invoke(null, inputs);
    }

    /// <summary>
    ///     Sets up RuntimeHelpers delegates for test execution.
    /// </summary>
    private static void SetupDelegates()
    {
        // ToResultDelegate — converts ExcelValue/IExcelRange results to object[,]
        RuntimeHelpers.ToResultDelegate ??= result =>
        {
            if (result is ExcelValue ev)
            {
                return ev.ToResult();
            }

            if (result is IExcelRange range)
            {
                return range.ToResult();
            }

            if (result is bool b)
            {
                return b.ToResult();
            }

            if (result is int i)
            {
                return i.ToResult();
            }

            if (result is double d)
            {
                return d.ToResult();
            }

            if (result is string s)
            {
                return s.ToResult();
            }

            // Handle LINQ IEnumerable<Row> results (from .Rows.Where(), .Rows.Select(), etc.)
            if (result is IEnumerable<Row> rows)
            {
                var rowList = rows.ToList();
                if (rowList.Count == 0)
                {
                    return string.Empty;
                }

                var cols = rowList[0].ColumnCount;
                var arr = new object?[rowList.Count, cols];
                for (var r = 0; r < rowList.Count; r++)
                for (var c = 0; c < cols; c++)
                {
                    arr[r, c] = rowList[r][c].Value;
                }

                return arr;
            }

            // Handle IEnumerable<ColumnValue> results (from .Rows.Select(r => r[0]))
            if (result is IEnumerable<ColumnValue> colValues)
            {
                var list = colValues.ToList();
                if (list.Count == 0)
                {
                    return string.Empty;
                }

                var arr = new object?[list.Count, 1];
                for (var r = 0; r < list.Count; r++)
                {
                    arr[r, 0] = list[r].Value;
                }

                return arr;
            }

            // Handle generic IEnumerable (other LINQ results)
            if (result is IEnumerable enumerable and not string and not object[,])
            {
                var list = enumerable.Cast<object>().ToList();
                if (list.Count == 0)
                {
                    return string.Empty;
                }

                var arr = new object?[list.Count, 1];
                for (var r = 0; r < list.Count; r++)
                {
                    arr[r, 0] = list[r] is ColumnValue cv ? cv.Value : list[r];
                }

                return arr;
            }

            return result ?? string.Empty;
        };

        // GetHeadersDelegate — for tests, extract first row as headers
        RuntimeHelpers.GetHeadersDelegate ??= rangeRef =>
        {
            if (rangeRef is not object[,] values)
            {
                return null;
            }

            var cols = values.GetLength(1);
            var headers = new string[cols];
            for (var i = 0; i < cols; i++)
            {
                headers[i] = values[0, i]?.ToString() ?? "";
            }

            return headers;
        };
    }

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
            // RuntimeHelpers (in formula-boss.dll)
            MetadataReference.CreateFromFile(typeof(RuntimeHelpers).Assembly.Location),
            // FormulaBoss.Runtime types
            MetadataReference.CreateFromFile(typeof(ExcelValue).Assembly.Location)
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
            // Ignore
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
}

public class NewTestCompilationResult
{
    public bool Success { get; init; }
    public string? ErrorMessage { get; init; }
    public MethodInfo? Method { get; init; }
    public string? MethodName { get; init; }
    public bool RequiresObjectModel { get; init; }
    public string? SourceCode { get; init; }
    public DetectionResult? Detection { get; init; }

    public string GetDiagnostics()
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== New Pipeline Test Compilation Diagnostics ===");
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
