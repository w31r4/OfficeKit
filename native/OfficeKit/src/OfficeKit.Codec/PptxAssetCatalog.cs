using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns durable picture-bullet asset identity. Open Packaging Convention part
// names and relationship IDs intentionally remain outside the wire contract.
internal sealed class PptxAssetCatalog
{
    private const int MaxAssets = 1_024;
    private const int MaxAssetBytes = 16 * 1024 * 1024;
    private const int MaxMediaAssetBytes = 64 * 1024 * 1024;
    private const string PictureAssetPrefix = "asset/presentation/picture-bullet/";
    private const string MediaAssetPrefix = "asset/presentation/media/";
    private const string OleWorkbookAssetPrefix = "asset/presentation/ole-workbook/";
    private const string OleOfficePackageAssetPrefix = "asset/presentation/ole-office-package/";
    private const string SmartArtDefinitionAssetPrefix = "asset/presentation/smartart-definition/";
    private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private const string DocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    internal const string SmartArtDefinitionContentType = "application/vnd.officekit.smartart-definition+json";
    private readonly Dictionary<string, Asset> _assets = new(StringComparer.Ordinal);
    private readonly Dictionary<string, ImagePart> _partByAssetId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, MediaDataPart> _mediaPartByAssetId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Asset> _imported = new(StringComparer.Ordinal);
    private readonly Func<ImagePart, string?>? _validatedPartSha256;
    private readonly ulong _maxTotalBytes;
    private readonly ulong _maxMediaAssetBytes;
    private ulong _totalBytes;

    internal PptxAssetCatalog(
        IEnumerable<Asset>? assets,
        EffectiveCodecLimits limits,
        Func<ImagePart, string?>? validatedPartSha256 = null)
    {
        _validatedPartSha256 = validatedPartSha256;
        _maxTotalBytes = Math.Min(limits.MaxUncompressedBytes, (ulong)MaxAssets * MaxAssetBytes);
        _maxMediaAssetBytes = Math.Min(limits.MaxInputBytes, (ulong)MaxMediaAssetBytes);
        foreach (var asset in assets ?? []) AddRequested(asset);
    }

    internal IReadOnlyCollection<Asset> ImportedAssets => _imported.Values;

    internal Asset Get(string assetId) => _assets.TryGetValue(assetId, out var asset)
        ? asset.Id.StartsWith(PictureAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation picture bullet references non-image asset {assetId}.")
        : throw new CodecException("invalid_presentation_asset", $"Presentation picture bullet references missing asset {assetId}.");

    internal Asset GetOleWorkbook(string assetId) => _assets.TryGetValue(assetId, out var asset) &&
        asset.Id.StartsWith(OleWorkbookAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation OLE workbook references missing asset {assetId}.");

    internal Asset GetOleOfficePackage(string assetId) => _assets.TryGetValue(assetId, out var asset) &&
        asset.Id.StartsWith(OleOfficePackageAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation OLE Office package references missing asset {assetId}.");

    internal Asset GetSmartArtDefinition(string assetId) => _assets.TryGetValue(assetId, out var asset) &&
        asset.Id.StartsWith(SmartArtDefinitionAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation SmartArt references missing definition asset {assetId}.");

    internal Asset GetMedia(string assetId) => _assets.TryGetValue(assetId, out var asset) &&
        asset.Id.StartsWith(MediaAssetPrefix, StringComparison.Ordinal)
            ? asset
            : throw new CodecException("invalid_presentation_asset", $"Presentation media references missing asset {assetId}.");

    internal MediaDataPart GetOrCreateMediaPart(PresentationDocument package, string assetId)
    {
        if (_mediaPartByAssetId.TryGetValue(assetId, out var existing)) return existing;
        var asset = GetMedia(assetId);
        var part = package.CreateMediaDataPart(asset.ContentType, MediaExtension(asset.ContentType));
        using (var source = new MemoryStream(asset.Data.ToByteArray(), writable: false)) part.FeedData(source);
        _mediaPartByAssetId.Add(assetId, part);
        return part;
    }

    internal Asset Import(ImagePart part)
    {
        var contentType = NormalizeContentType(part.ContentType);
        var digest = _validatedPartSha256?.Invoke(part);
        if (digest is { Length: 64 } && digest.All(char.IsAsciiHexDigit))
        {
            digest = digest.ToLowerInvariant();
            var validatedId = PictureAssetPrefix + digest;
            if (_assets.TryGetValue(validatedId, out var requestedAsset) &&
                requestedAsset.ContentType.Equals(contentType, StringComparison.OrdinalIgnoreCase))
            {
                _partByAssetId.TryAdd(validatedId, part);
                return requestedAsset;
            }
            if (_imported.TryGetValue(validatedId, out var importedAsset) &&
                importedAsset.ContentType.Equals(contentType, StringComparison.OrdinalIgnoreCase))
            {
                _partByAssetId.TryAdd(validatedId, part);
                return importedAsset;
            }
        }
        else
        {
            digest = null;
        }
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var data = memory.GetBuffer();
        var dataLength = checked((int)memory.Length);
        ValidateImage(contentType, data, dataLength, $"Presentation image part {part.Uri}");
        digest ??= Hash(data.AsSpan(0, dataLength));
        var id = PictureAssetPrefix + digest;
        if (_assets.TryGetValue(id, out var requested))
        {
            _partByAssetId.TryAdd(id, part);
            return requested;
        }
        if (!_imported.TryGetValue(id, out var asset))
        {
            if (_imported.Count >= MaxAssets)
                throw new CodecException("presentation_asset_budget_exceeded", $"Presentation exceeds the {MaxAssets}-asset budget.");
            EnsureBudget(dataLength);
            asset = new Asset
            {
                Id = id,
                FileName = $"picture-bullet-{digest[..16]}.{Extension(contentType)}",
                ContentType = contentType,
                Data = ByteString.CopyFrom(data, 0, dataLength),
                Sha256 = digest,
            };
            _imported.Add(id, asset);
        }
        _partByAssetId.TryAdd(id, part);
        return asset;
    }

    internal Asset ImportSmartArtDefinition(ReadOnlySpan<byte> data, string declaredSha256)
    {
        if (data.Length is < 1 or > 1024 * 1024)
            throw new CodecException("invalid_presentation_asset", "Presentation SmartArt definition must contain 1 through 1048576 bytes.");
        var digest = Hash(data);
        if (!string.IsNullOrWhiteSpace(declaredSha256) &&
            !declaredSha256.Equals(digest, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("invalid_presentation_asset", "Presentation SmartArt definition does not match its declared SHA-256 digest.");
        var id = SmartArtDefinitionAssetPrefix + digest;
        if (_assets.TryGetValue(id, out var requested)) return requested;
        if (_imported.TryGetValue(id, out var imported)) return imported;
        EnsureBudget(data.Length, 1024 * 1024, "SmartArt definition");
        var asset = new Asset
        {
            Id = id,
            FileName = $"smartart-definition-{digest[..16]}.json",
            ContentType = SmartArtDefinitionContentType,
            Data = ByteString.CopyFrom(data),
            Sha256 = digest,
        };
        _imported.Add(id, asset);
        return asset;
    }

    internal ImagePart? ExistingPart(string assetId) => _partByAssetId.GetValueOrDefault(assetId);

    internal void RegisterPart(string assetId, ImagePart part) => _partByAssetId.TryAdd(assetId, part);

    internal void IndexExistingParts(IEnumerable<ImagePart> parts)
    {
        foreach (var part in parts.Distinct())
        {
            try
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var memory = new MemoryStream();
                stream.CopyTo(memory);
                var data = memory.ToArray();
                var id = PictureAssetPrefix + Hash(data);
                if (_assets.TryGetValue(id, out var asset) && part.ContentType.Equals(asset.ContentType, StringComparison.OrdinalIgnoreCase))
                    _partByAssetId.TryAdd(id, part);
            }
            catch (IOException)
            {
                // Opaque/unreadable image parts are guarded elsewhere and are
                // never selected as modeled assets.
            }
        }
    }

    internal static PartTypeInfo ImagePartTypeFor(string contentType) => NormalizeContentType(contentType) switch
    {
        "image/png" => ImagePartType.Png,
        "image/jpeg" => ImagePartType.Jpeg,
        "image/gif" => ImagePartType.Gif,
        "image/svg+xml" => ImagePartType.Svg,
        _ => throw new CodecException("invalid_presentation_asset", $"Unsupported presentation image content type {contentType}."),
    };

    internal static string NativeAssetIdFor(string contentType, string sha256)
    {
        var normalized = NormalizeContentType(contentType);
        if (normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            return PictureAssetPrefix + sha256.ToLowerInvariant();
        if (normalized.StartsWith("audio/", StringComparison.OrdinalIgnoreCase) ||
            normalized.StartsWith("video/", StringComparison.OrdinalIgnoreCase))
            return MediaAssetPrefix + sha256.ToLowerInvariant();
        if (normalized.Equals(SpreadsheetContentType, StringComparison.Ordinal))
            return OleWorkbookAssetPrefix + sha256.ToLowerInvariant();
        if (normalized.Equals(DocumentContentType, StringComparison.Ordinal))
            return OleOfficePackageAssetPrefix + sha256.ToLowerInvariant();
        if (normalized.Equals(SmartArtDefinitionContentType, StringComparison.Ordinal))
            return SmartArtDefinitionAssetPrefix + sha256.ToLowerInvariant();
        throw new CodecException(
            "ppj.asset.unsupportedPurpose",
            $"PPJ presentation asset MIME {contentType} does not have a native compiler purpose.");
    }

    internal static bool IsCompilerOnlyAsset(Asset asset) =>
        NormalizeContentType(asset.ContentType).Equals(SmartArtDefinitionContentType, StringComparison.Ordinal);

    private void AddRequested(Asset source)
    {
        if (_assets.Count >= MaxAssets)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation exceeds the {MaxAssets}-asset budget.");
        var contentType = NormalizeContentType(source.ContentType);
        var data = source.Data.ToByteArray();
        var isPicture = source.Id.StartsWith(PictureAssetPrefix, StringComparison.Ordinal);
        var isMedia = source.Id.StartsWith(MediaAssetPrefix, StringComparison.Ordinal);
        var isOleWorkbook = source.Id.StartsWith(OleWorkbookAssetPrefix, StringComparison.Ordinal);
        var isOleOfficePackage = source.Id.StartsWith(OleOfficePackageAssetPrefix, StringComparison.Ordinal);
        var isSmartArtDefinition = source.Id.StartsWith(SmartArtDefinitionAssetPrefix, StringComparison.Ordinal);
        if (isPicture) ValidateImage(contentType, data, $"Presentation asset {source.Id}");
        else if (isMedia) ValidateMedia(contentType, data, $"Presentation asset {source.Id}");
        else if (isOleWorkbook) ValidateOleWorkbook(contentType, data, $"Presentation asset {source.Id}");
        else if (isOleOfficePackage) ValidateOleOfficePackage(contentType, data, $"Presentation asset {source.Id}");
        else if (isSmartArtDefinition)
        {
            if (!contentType.Equals(SmartArtDefinitionContentType, StringComparison.Ordinal) || data.Length is < 1 or > 1024 * 1024)
                throw new CodecException("invalid_presentation_asset", $"Presentation SmartArt definition asset {source.Id} has invalid bytes or MIME type.");
        }
        else throw new CodecException("invalid_presentation_asset", $"Presentation asset ID {source.Id} has an unsupported purpose prefix.");
        var digest = Hash(data);
        if (!source.Sha256.Equals(digest, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("invalid_presentation_asset", $"Presentation asset {source.Id} does not match its SHA-256 digest.");
        var expectedId = (isPicture
            ? PictureAssetPrefix
            : isMedia
                ? MediaAssetPrefix
                : isOleWorkbook
                    ? OleWorkbookAssetPrefix
                    : isOleOfficePackage
                        ? OleOfficePackageAssetPrefix
                        : SmartArtDefinitionAssetPrefix) + digest;
        if (!source.Id.Equals(expectedId, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_asset", $"Presentation asset {source.Id} is not content-addressed by its bytes.");
        if (!_assets.TryAdd(source.Id, source.Clone()))
            throw new CodecException("invalid_presentation_asset", $"Presentation contains duplicate asset ID {source.Id}.");
        EnsureBudget(
            data.LongLength,
            isMedia ? _maxMediaAssetBytes : isSmartArtDefinition ? 1024UL * 1024 : MaxAssetBytes,
            isMedia ? "media" : isSmartArtDefinition ? "SmartArt definition" : "image/OLE");
    }

    private void EnsureBudget(long length) => EnsureBudget(length, MaxAssetBytes, "picture-bullet");

    private void EnsureBudget(long length, ulong maximum, string purpose)
    {
        if (length <= 0 || (ulong)length > maximum)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation {purpose} assets must contain 1 through {maximum} bytes.");
        _totalBytes = checked(_totalBytes + (ulong)length);
        if (_totalBytes > _maxTotalBytes)
            throw new CodecException("presentation_asset_budget_exceeded", $"Presentation picture-bullet assets exceed the {_maxTotalBytes}-byte budget.");
    }

    private static void ValidateImage(string contentType, byte[] data, string label) =>
        ValidateImage(contentType, data, data.Length, label);

    private static void ValidateImage(string contentType, byte[] data, int dataLength, string label)
    {
        if (dataLength is 0 or > MaxAssetBytes)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxAssetBytes} bytes.");
        var bytes = data.AsSpan(0, dataLength);
        var valid = contentType switch
        {
            "image/png" => bytes.StartsWith(Convert.FromHexString("89504E470D0A1A0A")),
            "image/jpeg" => dataLength >= 3 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff,
            "image/gif" => dataLength >= 6 && Encoding.ASCII.GetString(data, 0, 6) is "GIF87a" or "GIF89a",
            "image/svg+xml" => IsSafeSvg(data, dataLength),
            _ => false,
        };
        if (!valid) throw new CodecException("invalid_presentation_asset", $"{label} bytes do not match a supported PNG, JPEG, GIF, or safe SVG content type.");
    }

    private static void ValidateOleWorkbook(string contentType, byte[] data, string label)
    {
        if (!contentType.Equals(SpreadsheetContentType, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_asset", $"{label} must use the XLSX workbook content type.");
        if (data.Length is 0 or > MaxAssetBytes || data.Length < 4 ||
            data[0] != 0x50 || data[1] != 0x4b || data[2] != 0x03 || data[3] != 0x04)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxAssetBytes} bytes of an OPC ZIP package.");
    }

    private static void ValidateOleOfficePackage(string contentType, byte[] data, string label)
    {
        if (!contentType.Equals(DocumentContentType, StringComparison.Ordinal))
            throw new CodecException("invalid_presentation_asset", $"{label} must use one supported Office package content type.");
        if (data.Length is 0 or > MaxAssetBytes || data.Length < 4 ||
            data[0] != 0x50 || data[1] != 0x4b || data[2] != 0x03 || data[3] != 0x04)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxAssetBytes} bytes of an OPC ZIP package.");
    }

    private static void ValidateMedia(string contentType, byte[] data, string label)
    {
        if (data.Length is 0 or > MaxMediaAssetBytes)
            throw new CodecException("invalid_presentation_asset", $"{label} must contain 1 through {MaxMediaAssetBytes} bytes.");
        var bytes = data.AsSpan();
        var isoBaseMedia = bytes.Length >= 12 && bytes.Slice(4, 4).SequenceEqual("ftyp"u8);
        var mpegAudio = bytes.StartsWith("ID3"u8) || bytes.Length >= 2 && data[0] == 0xff && (data[1] & 0xe0) == 0xe0;
        var wave = bytes.Length >= 12 && bytes.StartsWith("RIFF"u8) && bytes.Slice(8, 4).SequenceEqual("WAVE"u8);
        var valid = contentType switch
        {
            "video/mp4" or "audio/mp4" => isoBaseMedia,
            "audio/mpeg" => mpegAudio,
            "audio/wav" or "audio/x-wav" => wave,
            _ => false,
        };
        if (!valid)
            throw new CodecException("invalid_presentation_asset", $"{label} bytes do not match a supported MP4, MP3, M4A, or WAV content type.");
    }

    private static bool IsSafeSvg(byte[] data, int dataLength)
    {
        try
        {
            using var stream = new MemoryStream(data, 0, dataLength, writable: false, publiclyVisible: true);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaxAssetBytes,
                IgnoreComments = true,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root?.Name.LocalName != "svg") return false;
            if (document.DescendantNodes().OfType<XProcessingInstruction>().Any()) return false;
            foreach (var element in document.Root.DescendantsAndSelf())
            {
                if (element.Name.LocalName is "script" or "foreignObject") return false;
                if (element.Name.LocalName == "style" && UnsafeCss(element.Value)) return false;
                foreach (var attribute in element.Attributes())
                {
                    if (attribute.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase)) return false;
                    if (UnsafeCss(attribute.Value)) return false;
                    if (attribute.Name.LocalName != "href") continue;
                    var target = attribute.Value.Trim();
                    if (target.Length > 0 && !target.StartsWith('#') &&
                        !target.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase) &&
                        !target.StartsWith("data:image/jpeg;base64,", StringComparison.OrdinalIgnoreCase) &&
                        !target.StartsWith("data:image/gif;base64,", StringComparison.OrdinalIgnoreCase)) return false;
                }
            }
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static bool UnsafeCss(string value)
    {
        if (value.Contains("@import", StringComparison.OrdinalIgnoreCase)) return true;
        var offset = 0;
        while (value.IndexOf("url(", offset, StringComparison.OrdinalIgnoreCase) is var start && start >= 0)
        {
            var end = value.IndexOf(')', start + 4);
            if (end < 0) return true;
            var target = value[(start + 4)..end].Trim().Trim('\'', '"');
            if (target.Length > 0 && !target.StartsWith('#') &&
                !target.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase) &&
                !target.StartsWith("data:image/jpeg;base64,", StringComparison.OrdinalIgnoreCase) &&
                !target.StartsWith("data:image/gif;base64,", StringComparison.OrdinalIgnoreCase)) return true;
            offset = end + 1;
        }
        return false;
    }

    private static string NormalizeContentType(string value) => value.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)
        ? "image/jpeg"
        : value.ToLowerInvariant();

    private static string Extension(string contentType) => contentType switch
    {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/svg+xml" => "svg",
        _ => "bin",
    };

    private static string MediaExtension(string contentType) => contentType switch
    {
        "video/mp4" => ".mp4",
        "audio/mpeg" => ".mp3",
        "audio/mp4" => ".m4a",
        "audio/wav" or "audio/x-wav" => ".wav",
        _ => throw new CodecException("invalid_presentation_asset", $"Unsupported presentation media content type {contentType}."),
    };

    private static string Hash(byte[] data) => Hash(data.AsSpan());
    private static string Hash(ReadOnlySpan<byte> data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
}
