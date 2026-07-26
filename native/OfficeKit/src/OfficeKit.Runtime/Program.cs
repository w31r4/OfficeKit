using System.Runtime.Versioning;
using OfficeKit.Codec;

return 0;

[SupportedOSPlatform("browser")]
public partial class OfficeKitExports
{
    internal static byte[] Invoke(byte[] requestBytes) => CodecProtocol.Invoke(requestBytes);
}
