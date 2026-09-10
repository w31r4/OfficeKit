using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphSpaceAfterLifecyclePreservesSource(string kind) =>
        CheckParagraphSpacingLifecycle(kind, "spaceAfter");

    [Fact]
    public void ParagraphSpaceAfterStylePrecedenceSelectsUnit() =>
        CheckParagraphSpacingPrecedence("spaceAfter");

    [Fact]
    public void ParagraphSpaceAfterRetainsUnmodeledSource() =>
        CheckParagraphSpacingUnmodeledSource("spaceAfter");
}
