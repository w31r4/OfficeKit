using Xunit;

namespace OfficeKit.Codec.Tests;

public sealed partial class PpjTextBodyPropertyLifecycleTests
{
    [Theory]
    [InlineData("text")]
    [InlineData("shape")]
    public void ParagraphHangingLifecyclePreservesSource(string kind) =>
        CheckParagraphCoordinateLifecycle(kind, "hanging");

    [Fact]
    public void ParagraphHangingRetainsUnmodeledSource() =>
        CheckParagraphCoordinateUnmodeledSource("hanging");
}
