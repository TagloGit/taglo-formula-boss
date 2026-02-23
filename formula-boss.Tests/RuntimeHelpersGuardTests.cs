using System.Text.RegularExpressions;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Guards the critical invariant that RuntimeHelpers.cs must never reference ExcelDNA types.
///     A regression here causes TypeLoadException at runtime when generated code loads the class.
///     See CLAUDE.md "Delegate Bridge Pattern" for details.
/// </summary>
public class RuntimeHelpersGuardTests
{
    [Fact]
    public void RuntimeHelpers_HasNoExcelDnaUsingDirective()
    {
        var sourceFile = FindRuntimeHelpersSource();
        Assert.True(File.Exists(sourceFile), $"Could not find RuntimeHelpers.cs at {sourceFile}");

        var lines = File.ReadAllLines(sourceFile);

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            // Check actual using directives (not comments or strings)
            if (trimmed.StartsWith("using ") && !trimmed.StartsWith("using System"))
            {
                Assert.DoesNotContain("ExcelDna", trimmed);
            }
        }
    }

    [Fact]
    public void RuntimeHelpers_CompiledType_HasNoExcelDnaDependencies()
    {
        var type = typeof(RuntimeHelpers);

        // Check fields
        foreach (var field in type.GetFields(System.Reflection.BindingFlags.Public |
                                             System.Reflection.BindingFlags.NonPublic |
                                             System.Reflection.BindingFlags.Static |
                                             System.Reflection.BindingFlags.Instance))
        {
            Assert.DoesNotContain("ExcelDna", field.FieldType.FullName ?? "");
        }

        // Check method parameters and return types
        foreach (var method in type.GetMethods(System.Reflection.BindingFlags.Public |
                                               System.Reflection.BindingFlags.NonPublic |
                                               System.Reflection.BindingFlags.Static |
                                               System.Reflection.BindingFlags.Instance |
                                               System.Reflection.BindingFlags.DeclaredOnly))
        {
            Assert.DoesNotContain("ExcelDna", method.ReturnType.FullName ?? "");
            foreach (var param in method.GetParameters())
            {
                Assert.DoesNotContain("ExcelDna", param.ParameterType.FullName ?? "");
            }
        }

        // Check properties
        foreach (var prop in type.GetProperties(System.Reflection.BindingFlags.Public |
                                                System.Reflection.BindingFlags.NonPublic |
                                                System.Reflection.BindingFlags.Static |
                                                System.Reflection.BindingFlags.Instance))
        {
            Assert.DoesNotContain("ExcelDna", prop.PropertyType.FullName ?? "");
        }
    }

    private static string FindRuntimeHelpersSource()
    {
        var dir = AppContext.BaseDirectory;
        while (dir != null)
        {
            var candidate = Path.Combine(dir, "formula-boss", "RuntimeHelpers.cs");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            dir = Path.GetDirectoryName(dir);
        }

        return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "formula-boss",
            "RuntimeHelpers.cs"));
    }
}
