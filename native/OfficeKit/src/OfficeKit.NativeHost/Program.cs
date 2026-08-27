using System.Buffers.Binary;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using OfficeKit.Codec;

const int TransportVersion = 1;
const int HandshakeBytes = 12;
const int FramePrefixBytes = 4;
const int AbsoluteFrameLimit = 128 * 1024 * 1024;
const int LargeExchangeCollectionThreshold = 8 * 1024 * 1024;

if (args.Length == 1 && args[0] == "--status")
{
    Console.Out.WriteLine($"{{\"assemblyName\":\"officekit-codec\",\"backend\":\"native-aot\",\"transportVersion\":{TransportVersion},\"protocolVersion\":{CodecProtocol.ProtocolVersion}}}");
    return 0;
}
if (args.Length > 1 || args.Length == 1 && args[0] != "--serve")
{
    Console.Error.WriteLine("Usage: officekit-codec [--serve|--status]");
    return 64;
}

var input = Console.OpenStandardInput();
var output = Console.OpenStandardOutput();
var handshake = new byte[HandshakeBytes];
"OKIT"u8.CopyTo(handshake);
BinaryPrimitives.WriteInt32BigEndian(handshake.AsSpan(4, 4), TransportVersion);
BinaryPrimitives.WriteUInt32BigEndian(handshake.AsSpan(8, 4), CodecProtocol.ProtocolVersion);
await output.WriteAsync(handshake);
await output.FlushAsync();

var prefix = new byte[FramePrefixBytes];
while (true)
{
    var prefixRead = await ReadExactlyOrEofAsync(input, prefix);
    if (prefixRead == 0) return 0;
    if (prefixRead != prefix.Length)
    {
        Console.Error.WriteLine("OfficeKit native transport received a truncated frame prefix.");
        return 65;
    }

    var requestLength = BinaryPrimitives.ReadInt32BigEndian(prefix);
    if (requestLength <= 0 || requestLength > AbsoluteFrameLimit)
    {
        Console.Error.WriteLine("OfficeKit native transport rejected an invalid request frame length.");
        return 65;
    }

    var request = GC.AllocateUninitializedArray<byte>(requestLength);
    if (await ReadExactlyOrEofAsync(input, request) != requestLength)
    {
        Console.Error.WriteLine("OfficeKit native transport received a truncated request frame.");
        return 65;
    }

    var response = CodecProtocol.Invoke(ref request);
    if (response.Length > AbsoluteFrameLimit)
    {
        response = TransportFailure(
            "response_budget_exceeded",
            $"Codec response exceeds the absolute {AbsoluteFrameLimit}-byte native transport budget.");
    }
    var largeExchange = (long)requestLength + response.Length >= LargeExchangeCollectionThreshold;
    if (largeExchange)
    {
        // CodecProtocol has already serialized the response. Release the input
        // and its transient Open XML/protobuf graph before JavaScript expands
        // that response into the public object model.
        request = [];
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
    }
    BinaryPrimitives.WriteInt32BigEndian(prefix, response.Length);
    await output.WriteAsync(prefix);
    await output.WriteAsync(response);
    await output.FlushAsync();

    if (largeExchange)
    {
        // The pipe owns its copy now; do not retain the serialized payload in
        // the idle codec process either.
        response = [];
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
    }
}

static async ValueTask<int> ReadExactlyOrEofAsync(Stream stream, Memory<byte> buffer)
{
    var total = 0;
    while (total < buffer.Length)
    {
        var read = await stream.ReadAsync(buffer[total..]);
        if (read == 0) break;
        total += read;
    }
    return total;
}

static byte[] TransportFailure(string code, string message)
{
    var response = new CodecResponse
    {
        ProtocolVersion = CodecProtocol.ProtocolVersion,
        Ok = false,
    };
    response.Diagnostics.Add(new Diagnostic
    {
        Severity = DiagnosticSeverity.Error,
        Code = code,
        Message = message,
    });
    return response.ToByteArray();
}
