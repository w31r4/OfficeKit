using OfficeKit.Codec;
using OfficeKit.NativeHost;
using OfficeKit.PpjNativeHost;

if (args.Length > 0 && args[0] == "--build")
    return PpjNativeBuildCommand.Run(args[1..]);

return await NativeHostRunner.RunAsync(args, "officekit-ppj-codec", PpjCodecProtocol.InvokeResponse);
