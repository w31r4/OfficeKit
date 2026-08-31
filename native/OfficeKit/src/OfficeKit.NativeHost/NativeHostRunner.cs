using System.Buffers.Binary;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using OfficeKit.Codec;

namespace OfficeKit.NativeHost;

internal delegate CodecResponse CodecInvoke(ref byte[] request, byte[]? requestFile);

internal static class NativeHostRunner
{
    private const int TransportVersion = 2;
    private const int HandshakeBytes = 12;
    private const int RequestPrefixBytes = 8;
    private const int ResponsePrefixBytes = 4;
    private const int AbsoluteFrameLimit = 128 * 1024 * 1024;
    private const int LargeExchangeCollectionThreshold = 8 * 1024 * 1024;

    internal static async Task<int> RunAsync(string[] args, string assemblyName, CodecInvoke invoke)
    {
        if (args.Length == 1 && args[0] == "--status")
        {
            Console.Out.WriteLine($"{{\"assemblyName\":\"{assemblyName}\",\"backend\":\"native-aot\",\"transportVersion\":{TransportVersion},\"protocolVersion\":{CodecProtocol.ProtocolVersion}}}");
            return 0;
        }
        if (args.Length > 1 || args.Length == 1 && args[0] != "--serve")
        {
            Console.Error.WriteLine($"Usage: {assemblyName} [--serve|--status]");
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

        var prefix = new byte[RequestPrefixBytes];
        while (true)
        {
            var prefixRead = await ReadExactlyOrEofAsync(input, prefix);
            if (prefixRead == 0) return 0;
            if (prefixRead != prefix.Length)
            {
                Console.Error.WriteLine("OfficeKit native transport received a truncated frame prefix.");
                return 65;
            }

            var requestLength = BinaryPrimitives.ReadInt32BigEndian(prefix.AsSpan(0, 4));
            var requestFileLength = BinaryPrimitives.ReadInt32BigEndian(prefix.AsSpan(4, 4));
            if (requestLength <= 0 || requestFileLength < 0 || (long)requestLength + requestFileLength > AbsoluteFrameLimit)
            {
                Console.Error.WriteLine("OfficeKit native transport rejected an invalid request or file-sidecar frame length.");
                return 65;
            }

            var request = GC.AllocateUninitializedArray<byte>(requestLength);
            if (await ReadExactlyOrEofAsync(input, request) != requestLength)
            {
                Console.Error.WriteLine("OfficeKit native transport received a truncated request frame.");
                return 65;
            }

            var requestFile = requestFileLength == 0 ? [] : GC.AllocateUninitializedArray<byte>(requestFileLength);
            if (requestFileLength > 0 && await ReadExactlyOrEofAsync(input, requestFile) != requestFileLength)
            {
                Console.Error.WriteLine("OfficeKit native transport received a truncated file sidecar.");
                return 65;
            }

            CodecResponse? response = invoke(ref request, requestFile);
            requestFile = [];
            var responseLength = response.CalculateSize();
            if (responseLength > AbsoluteFrameLimit)
            {
                response = TransportFailure(
                    "response_budget_exceeded",
                    $"Codec response exceeds the absolute {AbsoluteFrameLimit}-byte native transport budget.");
                responseLength = response.CalculateSize();
            }
            var largeExchange = (long)requestLength + requestFileLength + responseLength >= LargeExchangeCollectionThreshold;
            if (largeExchange)
            {
                request = [];
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
            }
            BinaryPrimitives.WriteInt32BigEndian(prefix, responseLength);
            await output.WriteAsync(prefix.AsMemory(0, ResponsePrefixBytes));
            response.WriteTo(output);
            await output.FlushAsync();

            if (largeExchange)
            {
                response = null;
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
            }
        }
    }

    private static async ValueTask<int> ReadExactlyOrEofAsync(Stream stream, Memory<byte> buffer)
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

    private static CodecResponse TransportFailure(string code, string message)
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
        return response;
    }
}
