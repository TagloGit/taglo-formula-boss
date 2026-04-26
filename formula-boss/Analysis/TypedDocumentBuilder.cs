using System.Reflection;
using System.Text;

using FormulaBoss.Runtime;
using FormulaBoss.Transpilation;
using FormulaBoss.UI;

namespace FormulaBoss.Analysis;

/// <summary>
///     Generates synthetic typed classes (Row, RowCollection, Table) and usings
///     for Roslyn compilation contexts. Shared between the editor (completions,
///     diagnostics, hover) and the pipeline (semantic analysis).
/// </summary>
internal static class TypedDocumentBuilder
{
    private static readonly NullabilityInfoContext NullabilityCtx = new();

    private static readonly Dictionary<Type, string> CSharpKeywords = new()
    {
        [typeof(object)] = "object",
        [typeof(int)] = "int",
        [typeof(bool)] = "bool",
        [typeof(string)] = "string",
        [typeof(void)] = "void",
        [typeof(double)] = "double",
        [typeof(float)] = "float",
        [typeof(long)] = "long",
        [typeof(decimal)] = "decimal"
    };

    /// <summary>
    ///     Appends standard usings for the synthetic compilation context.
    /// </summary>
    public static void AppendUsings(StringBuilder sb)
    {
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using System.Text;");
        sb.AppendLine("using System.Text.RegularExpressions;");
        sb.AppendLine("using FormulaBoss.Runtime;");
        sb.AppendLine();
    }

    /// <summary>
    ///     Emits typed Row, RowCollection, and Table classes for each table in the metadata,
    ///     and returns a dictionary mapping table names to their synthetic type names.
    /// </summary>
    /// <returns>
    ///     A case-insensitive dictionary mapping original table names to synthetic table type names
    ///     (e.g. "Sales" → "__SalesTable").
    /// </returns>
    public static Dictionary<string, string> EmitTableTypes(
        StringBuilder sb, WorkbookMetadata? metadata)
    {
        var tableTypeNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (metadata == null)
        {
            return tableTypeNames;
        }

        foreach (var tableName in metadata.TableColumns.Keys)
        {
            var safeTableName = ColumnMapper.Sanitise(tableName);
            if (string.IsNullOrEmpty(safeTableName))
            {
                continue;
            }

            var rowTypeName = $"__{safeTableName}Row";
            var rowCollTypeName = $"__{safeTableName}RowCollection";
            var rowGroupTypeName = $"__{safeTableName}RowGroup";
            var groupedCollTypeName = $"__{safeTableName}GroupedRowCollection";
            var tableTypeName = $"__{safeTableName}Table";
            tableTypeNames[tableName] = tableTypeName;

            var colCollTypeName = $"__{safeTableName}ColumnCollection";

            var typeMap = new Dictionary<Type, string>
            {
                [typeof(RowCollection)] = rowCollTypeName,
                [typeof(GroupedRowCollection)] = groupedCollTypeName,
                [typeof(Row)] = rowTypeName,
                [typeof(RowGroup)] = rowGroupTypeName,
                [typeof(ColumnCollection)] = colCollTypeName,
                [typeof(Column)] = nameof(Column)
            };

            EmitTypedRow(sb, rowTypeName);
            EmitSyntheticCollection(sb, typeof(RowCollection), rowCollTypeName, typeMap, rowTypeName);
            EmitSyntheticCollection(sb, typeof(RowGroup), rowGroupTypeName, typeMap, rowTypeName);
            EmitSyntheticCollection(sb, typeof(GroupedRowCollection), groupedCollTypeName, typeMap,
                rowGroupTypeName);
            EmitSyntheticCollection(sb, typeof(ColumnCollection), colCollTypeName, typeMap,
                nameof(Column));
            EmitTypedTable(sb, tableTypeName, rowCollTypeName, colCollTypeName);
        }

        return tableTypeNames;
    }

    private static void EmitTypedRow(StringBuilder sb, string rowTypeName)
    {
        // Inherit ExcelArray so synthetic Row exposes the same LINQ/ExcelArray surface
        // (Skip, Where, Map, FirstOrDefault, ...) as the runtime Row : ExcelArray type.
        // Column names are surfaced separately by CompletionHelpers.BuildRowCompletions
        // as bracket-inserting items, so they're deliberately not emitted as properties
        // on the synthetic Row.
        sb.AppendLine($"class {rowTypeName} : ExcelArray {{");
        sb.AppendLine($"{rowTypeName}() : base(new object[0,0]) {{}}");

        // Row-specific members not present on ExcelArray
        sb.AppendLine("public ExcelScalar this[string columnName] => default!;");
        sb.AppendLine("public int ColumnCount => 0;");

        sb.AppendLine("}");
        sb.AppendLine();
    }

    /// <summary>
    ///     Reflects over a runtime collection type annotated with <see cref="SyntheticCollectionAttribute" />
    ///     and emits a synthetic stub class with type-substituted method signatures.
    /// </summary>
    private static void EmitSyntheticCollection(
        StringBuilder sb, Type runtimeType, string syntheticTypeName,
        Dictionary<Type, string> typeMap, string syntheticElementName)
    {
        var collAttr = runtimeType.GetCustomAttribute<SyntheticCollectionAttribute>();
        if (collAttr == null)
        {
            return;
        }

        // Find IEnumerable<T> to determine the enumerable element type
        var enumerableElementName = syntheticElementName;
        var enumerableInterface = runtimeType.GetInterfaces()
            .FirstOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));
        if (enumerableInterface != null)
        {
            var enumerableArg = enumerableInterface.GetGenericArguments()[0];
            if (typeMap.TryGetValue(enumerableArg, out var mappedName))
            {
                enumerableElementName = mappedName;
            }
        }

        sb.AppendLine($"class {syntheticTypeName} : IEnumerable<{enumerableElementName}> {{");

        // Collect [SyntheticMember] members (includes inherited)
        var members = runtimeType
            .GetMembers(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.GetCustomAttribute<SyntheticMemberAttribute>() != null)
            .ToList();

        foreach (var member in members)
        {
            switch (member)
            {
                case PropertyInfo prop:
                    EmitProperty(sb, prop, typeMap);
                    break;
                case MethodInfo method:
                    EmitMethod(sb, method, typeMap, syntheticElementName);
                    break;
            }
        }

        // IEnumerable boilerplate
        sb.AppendLine($"public IEnumerator<{enumerableElementName}> GetEnumerator() => default!;");
        sb.AppendLine("IEnumerator IEnumerable.GetEnumerator() => default!;");

        sb.AppendLine("}");
        sb.AppendLine();
    }

    private static void EmitProperty(
        StringBuilder sb, PropertyInfo prop, Dictionary<Type, string> typeMap)
    {
        var nullable = IsNullableProperty(prop);
        var typeName = FormatReturnType(prop.PropertyType, typeMap, nullable);
        var defaultExpr = GetDefaultExpression(prop.PropertyType, nullable);
        sb.AppendLine($"public {typeName} {prop.Name} => {defaultExpr};");
    }

    private static void EmitMethod(
        StringBuilder sb, MethodInfo method, Dictionary<Type, string> typeMap, string syntheticElementName)
    {
        var nullable = IsNullableReturn(method);
        var returnTypeName = FormatReturnType(method.ReturnType, typeMap, nullable);
        var parameters = method.GetParameters();
        var paramStrings = parameters
            .Select(p => FormatParameter(p, typeMap, syntheticElementName))
            .ToList();

        var paramList = string.Join(", ", paramStrings);
        var defaultExpr = GetDefaultExpression(method.ReturnType, nullable);
        sb.AppendLine($"public {returnTypeName} {method.Name}({paramList}) => {defaultExpr};");
    }

    private static string FormatParameter(
        ParameterInfo param, Dictionary<Type, string> typeMap, string syntheticElementName)
    {
        var type = param.ParameterType;
        string typeName;

        if (type.IsGenericType && IsFuncType(type))
        {
            typeName = FormatFuncType(param, typeMap, syntheticElementName);
        }
        else if (IsDynamicParameter(param))
        {
            typeName = "dynamic";
        }
        else if (typeMap.TryGetValue(type, out var mapped))
        {
            typeName = mapped;
        }
        else
        {
            typeName = FormatSimpleType(type);
        }

        return $"{typeName} {param.Name}";
    }

    private static string FormatFuncType(
        ParameterInfo param, Dictionary<Type, string> typeMap, string syntheticElementName)
    {
        var funcType = param.ParameterType;
        var typeArgs = funcType.GetGenericArguments();
        var dynamicFlags = GetDynamicFlags(param);
        var elementArgIndex = typeArgs.Length - 2; // Convention: element is at argCount - 2

        var formattedArgs = new string[typeArgs.Length];
        for (var i = 0; i < typeArgs.Length; i++)
        {
            var argType = typeArgs[i];
            // DynamicAttribute flag index: +1 because index 0 is for the Func type itself
            var isDynamic = dynamicFlags != null && i + 1 < dynamicFlags.Length && dynamicFlags[i + 1];

            if (isDynamic && i == elementArgIndex)
            {
                // This is the element type position — substitute
                formattedArgs[i] = syntheticElementName;
            }
            else if (isDynamic)
            {
                // Non-element dynamic position
                formattedArgs[i] = "dynamic";
            }
            else if (typeMap.TryGetValue(argType, out var mapped))
            {
                formattedArgs[i] = mapped;
            }
            else
            {
                formattedArgs[i] = FormatSimpleType(argType);
            }
        }

        return $"Func<{string.Join(", ", formattedArgs)}>";
    }

    private static string FormatReturnType(Type type, Dictionary<Type, string> typeMap, bool nullable)
    {
        if (typeMap.TryGetValue(type, out var mapped))
        {
            return nullable ? $"{mapped}?" : mapped;
        }

        if (type == typeof(object))
        {
            return nullable ? "object?" : "dynamic";
        }

        if (type.IsGenericType)
        {
            var def = type.GetGenericTypeDefinition();
            var args = type.GetGenericArguments();
            var baseName = def.Name[..def.Name.IndexOf('`')];
            var formattedArgs = args.Select(a => FormatReturnType(a, typeMap, false));
            return $"{baseName}<{string.Join(", ", formattedArgs)}>";
        }

        var name = FormatSimpleType(type);
        return nullable ? $"{name}?" : name;
    }

    private static string FormatSimpleType(Type type) =>
        CSharpKeywords.TryGetValue(type, out var keyword) ? keyword : type.Name;

    private static bool IsFuncType(Type type)
    {
        if (!type.IsGenericType)
        {
            return false;
        }

        var def = type.GetGenericTypeDefinition();
        return def.FullName?.StartsWith("System.Func`") == true;
    }

    private static bool[] GetDynamicFlags(ParameterInfo param)
    {
        var attr = param.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.FullName ==
                                 "System.Runtime.CompilerServices.DynamicAttribute");
        if (attr == null)
        {
            return Array.Empty<bool>();
        }

        if (attr.ConstructorArguments.Count == 0)
        {
            // Bare [Dynamic] — single true flag
            return new[] { true };
        }

        // [Dynamic(new[] { false, true, ... })]
        var flags = attr.ConstructorArguments[0].Value;
        if (flags is IReadOnlyCollection<CustomAttributeTypedArgument> args)
        {
            return args.Select(a => (bool)a.Value!).ToArray();
        }

        return Array.Empty<bool>();
    }

    private static bool IsDynamicParameter(ParameterInfo param)
    {
        var flags = GetDynamicFlags(param);
        return flags.Length > 0 && flags[0];
    }

    private static bool IsNullableReturn(MethodInfo method)
    {
        try
        {
            var info = NullabilityCtx.Create(method.ReturnParameter);
            return info.ReadState == NullabilityState.Nullable;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsNullableProperty(PropertyInfo prop)
    {
        try
        {
            var info = NullabilityCtx.Create(prop);
            return info.ReadState == NullabilityState.Nullable;
        }
        catch
        {
            return false;
        }
    }

    private static string GetDefaultExpression(Type returnType, bool nullable)
    {
        // Value types use 'default', reference types use 'default!'
        if (returnType.IsValueType && !nullable)
        {
            return "default";
        }

        // Nullable value type or reference type — for stubs, always default! for non-null ref,
        // default for nullable and value types
        if (nullable || (returnType.IsValueType && Nullable.GetUnderlyingType(returnType) != null))
        {
            return "default";
        }

        return "default!";
    }

    private static void EmitTypedTable(
        StringBuilder sb, string tableTypeName, string rowCollTypeName, string colCollTypeName)
    {
        sb.AppendLine($"class {tableTypeName} : ExcelTable {{");
        sb.AppendLine($"public new {rowCollTypeName} Rows => default!;");
        sb.AppendLine($"public new {colCollTypeName} Cols => default!;");
        sb.AppendLine("public new Column this[string columnName] => default!;");
        sb.AppendLine($"{tableTypeName}() : base(new object[0,0], System.Array.Empty<string>()) {{}}");
        sb.AppendLine("}");
        sb.AppendLine();
    }
}
