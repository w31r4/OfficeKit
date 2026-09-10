using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphLineSpacingLifecyclePreservesSource(string kind) =>
        CheckParagraphSpacingLifecycle(kind, "lineSpacing");

    [Fact]
    public void ParagraphLineSpacingStylePrecedenceSelectsUnit() =>
        CheckParagraphSpacingPrecedence("lineSpacing");

    [Fact]
    public void ParagraphLineSpacingRetainsUnmodeledSource() =>
        CheckParagraphSpacingUnmodeledSource("lineSpacing");
}
