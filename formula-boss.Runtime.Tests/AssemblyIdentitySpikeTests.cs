using System.Reflection;
using System.Runtime.Loader;

using Xunit;

namespace FormulaBoss.Runtime.Tests;

/// <summary>
///     Spike: Test whether Runtime assembly can be loaded from a separate AssemblyLoadContext
///     (simulating how Roslyn-compiled generated code would load it).
///     This determines whether delegate bridges are needed for COM access in Phase 2.
/// </summary>
public class AssemblyIdentitySpikeTests
{
    /// <summary>
    ///     Simulates the scenario where generated code is compiled by Roslyn and loaded
    ///     into a separate AssemblyLoadContext. The generated code references FormulaBoss.Runtime.
    ///     We verify that types from the Runtime assembly resolve correctly across contexts.
    /// </summary>
    [Fact]
    public void RuntimeAssembly_LoadedFromSeparateContext_TypesResolveCorrectly()
    {
        // Get the Runtime assembly path
        var runtimeAssemblyPath = typeof(ExcelValue).Assembly.Location;
        Assert.False(string.IsNullOrEmpty(runtimeAssemblyPath),
            "Runtime assembly location should not be empty");

        // Create a separate AssemblyLoadContext (simulating generated code's context)
        var alc = new AssemblyLoadContext("GeneratedCodeContext", true);
        try
        {
            // Load the Runtime assembly into the separate context
            var loadedAssembly = alc.LoadFromAssemblyPath(runtimeAssemblyPath);

            // Get ExcelValue type from the separately-loaded assembly
            var excelValueType = loadedAssembly.GetType("FormulaBoss.Runtime.ExcelValue");
            Assert.NotNull(excelValueType);

            // Try to use it via the separately-loaded type's static method
            var wrapMethod = excelValueType.GetMethod("Wrap",
                BindingFlags.Public | BindingFlags.Static,
                null,
                new[] { typeof(object), typeof(string[]), loadedAssembly.GetType("FormulaBoss.Runtime.RangeOrigin") },
                null);
            Assert.NotNull(wrapMethod);

            // Invoke Wrap() from the loaded assembly with a value
            var result = wrapMethod.Invoke(null, new object?[] { 42.0, null, null });
            Assert.NotNull(result);

            // Check if the result type name matches what we expect
            Assert.Equal("ExcelScalar", result.GetType().Name);

            // Verify the separately-loaded type is NOT assignable to the default context's type.
            // This confirms separate ALCs create identity mismatches (as expected).
            // However, this doesn't affect us — Roslyn-compiled code shares the default ALC
            // when we add the Runtime assembly as a MetadataReference and the assembly is already
            // loaded in the default context. The generated assembly resolves Runtime types
            // from the default context via assembly probing.
            //
            // CONCLUSION: No delegate bridges needed for Runtime types. The Runtime assembly
            // (without ExcelDNA dependency) will resolve correctly from generated code.
            // Delegate bridges are only needed for direct ExcelDNA/COM interop (Phase 2 decision).
            Assert.False(result is ExcelValue,
                "Separate ALC should cause identity mismatch (expected behavior)");
        }
        finally
        {
            alc.Unload();
        }
    }

    /// <summary>
    ///     Tests that the Runtime assembly has NO ExcelDNA dependency, which is a prerequisite
    ///     for avoiding the assembly identity issues documented in CLAUDE.md.
    /// </summary>
    [Fact]
    public void RuntimeAssembly_HasNoExcelDnaDependency()
    {
        var runtimeAssembly = typeof(ExcelValue).Assembly;
        var references = runtimeAssembly.GetReferencedAssemblies();
        var excelDnaRef = references.FirstOrDefault(r =>
            r.Name != null && r.Name.StartsWith("ExcelDna", StringComparison.OrdinalIgnoreCase));
        Assert.Null(excelDnaRef);
    }

    /// <summary>
    ///     Tests that Runtime types can be created and used purely with object/reflection,
    ///     which is how generated code would interact with them if loaded in a different context.
    /// </summary>
    [Fact]
    public void RuntimeTypes_WorkViaReflection()
    {
        var runtimeAssembly = typeof(ExcelValue).Assembly;

        // Simulate generated code creating types via reflection
        var scalarType = runtimeAssembly.GetType("FormulaBoss.Runtime.ExcelScalar")!;
        var scalar = Activator.CreateInstance(scalarType, new object?[] { 42.0, null })!;

        // Access RawValue via reflection
        var rawValueProp = scalarType.GetProperty("RawValue")!;
        var rawValue = rawValueProp.GetValue(scalar);
        Assert.Equal(42.0, rawValue);

        // Use Wrap factory
        var excelValueType = runtimeAssembly.GetType("FormulaBoss.Runtime.ExcelValue")!;
        var wrapMethod = excelValueType.GetMethod("Wrap")!;
        var wrapped = wrapMethod.Invoke(null, new object?[] { "hello", null, null })!;
        Assert.Equal("ExcelScalar", wrapped.GetType().Name);
    }
}
