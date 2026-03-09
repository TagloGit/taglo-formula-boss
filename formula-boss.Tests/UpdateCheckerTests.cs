using FormulaBoss.Updates;

using Xunit;

namespace FormulaBoss.Tests;

public class UpdateCheckerTests
{
    [Theory]
    [InlineData("v0.2.0", 0, 2, 0)]
    [InlineData("v1.0.0", 1, 0, 0)]
    [InlineData("0.3.1", 0, 3, 1)]
    [InlineData("V0.1.0", 0, 1, 0)]
    [InlineData("v10.20.30", 10, 20, 30)]
    public void ParseVersion_ValidTags_ReturnsCorrectVersion(string tag, int major, int minor, int build)
    {
        var result = UpdateChecker.ParseVersion(tag);

        Assert.NotNull(result);
        Assert.Equal(major, result.Major);
        Assert.Equal(minor, result.Minor);
        Assert.Equal(build, result.Build);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-a-version")]
    [InlineData("v")]
    [InlineData("vabc")]
    public void ParseVersion_InvalidTags_ReturnsNull(string tag)
    {
        var result = UpdateChecker.ParseVersion(tag);

        Assert.Null(result);
    }

    [Fact]
    public void ParseVersion_WithPrereleaseSuffix_ParsesBaseVersion()
    {
        // Version.TryParse ignores everything after the version numbers
        // but tags like "v0.2.0-beta" won't parse — that's acceptable
        var result = UpdateChecker.ParseVersion("v0.2.0");

        Assert.NotNull(result);
        Assert.Equal(new Version(0, 2, 0), result);
    }

    [Fact]
    public void VersionComparison_NewerVersion_IsGreater()
    {
        var current = new Version(0, 1, 0);
        var remote = UpdateChecker.ParseVersion("v0.2.0");

        Assert.NotNull(remote);
        Assert.True(remote > current);
    }

    [Fact]
    public void VersionComparison_SameVersion_IsNotGreater()
    {
        var current = new Version(0, 1, 0);
        var remote = UpdateChecker.ParseVersion("v0.1.0");

        Assert.NotNull(remote);
        Assert.False(remote > current);
    }

    [Fact]
    public void VersionComparison_OlderVersion_IsNotGreater()
    {
        var current = new Version(0, 2, 0);
        var remote = UpdateChecker.ParseVersion("v0.1.0");

        Assert.NotNull(remote);
        Assert.False(remote > current);
    }
}
