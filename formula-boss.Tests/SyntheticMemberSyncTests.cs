using System.Reflection;

using FormulaBoss.Runtime;

using Xunit;

namespace FormulaBoss.Tests;

/// <summary>
///     Verifies that every public instance method and property on runtime collection
///     classes has either <see cref="SyntheticMemberAttribute" /> or <see cref="SyntheticExcludeAttribute" />.
///     Catches forgotten annotations when new methods are added.
/// </summary>
public class SyntheticMemberSyncTests
{
    [Theory]
    [InlineData(typeof(RowCollection))]
    [InlineData(typeof(RowGroup))]
    [InlineData(typeof(GroupedRowCollection))]
    [InlineData(typeof(ColumnCollection))]
    public void AllPublicMembers_HaveSyntheticAttribute(Type collectionType)
    {
        var members = collectionType
            .GetMembers(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Where(m => m is MethodInfo { IsSpecialName: false } or PropertyInfo)
            .Where(m => m.DeclaringType != typeof(object))
            .ToList();

        var unmarked = members
            .Where(m => m.GetCustomAttribute<SyntheticMemberAttribute>() == null
                        && m.GetCustomAttribute<SyntheticExcludeAttribute>() == null)
            .Select(m => $"{collectionType.Name}.{m.Name}")
            .ToList();

        Assert.True(unmarked.Count == 0,
            $"Missing [SyntheticMember] or [SyntheticExclude] on: {string.Join(", ", unmarked)}");
    }

    [Theory]
    [InlineData(typeof(RowCollection))]
    [InlineData(typeof(RowGroup))]
    [InlineData(typeof(GroupedRowCollection))]
    [InlineData(typeof(ColumnCollection))]
    public void CollectionClass_HasSyntheticCollectionAttribute(Type collectionType)
    {
        var attr = collectionType.GetCustomAttribute<SyntheticCollectionAttribute>();
        Assert.NotNull(attr);
        Assert.NotNull(attr.ElementType);
    }
}
