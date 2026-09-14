using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using Microsoft.Win32.SafeHandles;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeKit.Codec;

internal sealed record PptxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics)
{
    // Read-only import provenance, allocated only by scene callers. Never
    // confused with source edit capabilities or serialized into the artifact.
    internal IReadOnlyList<PptxNativeBinding> NativeBindings { get; init; } = [];
}
internal sealed record PptxExportResult(
    byte[] File,
    IReadOnlyList<Diagnostic> Diagnostics,
    IReadOnlyList<string>? ChangedParts = null);
internal sealed record PptxNativeBinding(string PageId, string ElementId, string Type, string PartPath, uint NativeId);

internal sealed class PptxPackageSource : IDisposable
{
    private byte[]? bytes;
    private readonly FileStream? file;
    private readonly long length;
    private string? sha256;

    internal PptxPackageSource(byte[] bytes)
    {
        this.bytes = bytes;
        length = bytes.LongLength;
    }

    internal PptxPackageSource(string path)
    {
        file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        length = file.Length;
    }

    internal long Length => length;
    internal Stream OpenRead()
    {
        if (bytes is not null) return new MemoryStream(bytes, writable: false);
        var borrowedHandle = new SafeFileHandle(file!.SafeFileHandle.DangerousGetHandle(), ownsHandle: false);
        var stream = new FileStream(borrowedHandle, FileAccess.Read, bufferSize: 64 * 1024, isAsync: false);
        stream.Position = 0;
        return stream;
    }

    internal string Sha256()
    {
        if (sha256 is not null) return sha256;
        using var stream = OpenRead();
        sha256 = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        return sha256;
    }

    internal byte[] Materialize()
    {
        if (bytes is not null) return bytes;
        using var stream = OpenRead();
        var materialized = new byte[checked((int)length)];
        stream.ReadExactly(materialized);
        bytes = materialized;
        return bytes;
    }

    internal void CopyTo(Stream destination)
    {
        using var stream = OpenRead();
        stream.CopyTo(destination);
    }

    internal bool TryGetMaterialized(out byte[] value)
    {
        value = bytes!;
        return bytes is not null;
    }

    public void Dispose() => file?.Dispose();
}

internal interface IPptxSourceFreeBuildPlan
{
    PresentationArtifact Presentation { get; }
    bool RequiresPreviousSlide(int slideIndex);
    PresentationSlide MaterializeSlide(int slideIndex, PresentationSlide? previousSlide);
    void RecordNativeBindings(int slideIndex, PresentationSlide slide, IReadOnlyList<PresentationElement> flattenedElements);
}

internal sealed class PptxArtifactSourceFreeBuildPlan(PresentationArtifact presentation) : IPptxSourceFreeBuildPlan
{
    public PresentationArtifact Presentation { get; } = presentation;

    public bool RequiresPreviousSlide(int slideIndex) =>
        slideIndex >= 0 && slideIndex < Presentation.Slides.Count && Presentation.Slides[slideIndex].Morph is not null;

    public PresentationSlide MaterializeSlide(int slideIndex, PresentationSlide? previousSlide) =>
        Presentation.Slides[slideIndex];

    public void RecordNativeBindings(int slideIndex, PresentationSlide slide, IReadOnlyList<PresentationElement> flattenedElements)
    {
    }
}

internal sealed record PptxLayoutGraphEntry(int Index, string Id, string RelationshipId, SlideLayoutPart Part);
internal sealed record PptxMasterGraphEntry(int Index, string Id, string RelationshipId, SlideMasterPart Part, IReadOnlyList<PptxLayoutGraphEntry> Layouts);
internal sealed record PptxSourceSlideEntry(int Index, P.SlideId SlideId, string RelationshipId, SlidePart Part);
internal sealed class PptxTargetSlideEntry
{
    internal PptxTargetSlideEntry(int targetIndex, PresentationSlide target, PptxSourceSlideEntry source, bool isClone)
    {
        TargetIndex = targetIndex;
        Target = target;
        Source = source;
        IsClone = isClone;
        OutputSlideId = source.SlideId;
        OutputPart = source.Part;
    }

    internal int TargetIndex { get; }
    internal PresentationSlide Target { get; }
    internal PptxSourceSlideEntry Source { get; }
    internal bool IsClone { get; }
    internal P.SlideId OutputSlideId { get; set; }
    internal SlidePart OutputPart { get; set; }
}

internal static class PptxCodec
{
    private sealed class EnvelopeValidationContext(
        PptxAssetCatalog assetCatalog,
        bool hasSourcePackage,
        HashSet<string> layoutIds,
        Dictionary<string, PresentationMaster> mastersById,
        Dictionary<string, PresentationLayout> layoutsById,
        ulong items)
    {
        internal PptxAssetCatalog AssetCatalog { get; } = assetCatalog;
        internal bool HasSourcePackage { get; } = hasSourcePackage;
        internal HashSet<string> LayoutIds { get; } = layoutIds;
        internal Dictionary<string, PresentationMaster> MastersById { get; } = mastersById;
        internal Dictionary<string, PresentationLayout> LayoutsById { get; } = layoutsById;
        internal ulong Items { get; set; } = items;
    }

    internal static bool SupportsBoundTextLeaf(P.Shape shape) =>
        shape.TextBody is not null && PptxTextCodec.SupportsEditing(shape.TextBody);

    // A shape can have a perfectly usable local text graph even when its
    // visual fill/effect graph is outside the semantic projection (for
    // example, a gradient-filled banner).  Keep this narrow escape hatch
    // text-only: the native style, frame, geometry, accessibility and other
    // fields must remain byte-owned by the source package.
    private static bool IsBoundTextOnlyShapeEdit(
        PresentationElement original,
        PresentationElement requested,
        P.Shape sourceShape)
    {
        if (!string.Equals(original.Name, requested.Name, StringComparison.Ordinal) ||
            original.ContentCase != PresentationElement.ContentOneofCase.Shape ||
            requested.ContentCase != PresentationElement.ContentOneofCase.Shape ||
            original.Shape is null || requested.Shape is null ||
            !SupportsBoundTextLeaf(sourceShape))
            return false;
        var textOnly = original.Shape.Clone();
        textOnly.Text = requested.Shape.Text;
        textOnly.TextBody = requested.Shape.TextBody?.Clone();
        return textOnly.Equals(requested.Shape);
    }

    internal static int ValidateEditPlanOutput(
        byte[] sourceBytes,
        byte[] outputBytes,
        EffectiveCodecLimits limits)
    {
        ValidateOutputBudget(outputBytes, limits);
        return ValidateOffice2021AgainstSource(sourceBytes, outputBytes);
    }

    private const long DefaultSlideWidthEmu = 12_192_000;
    private const long DefaultSlideHeightEmu = 6_858_000;
    internal static PptxImportResult Import(
        byte[] bytes,
        EffectiveCodecLimits limits,
        bool retainImportedAssetData = true,
        string? verifiedPackageSha256 = null,
        bool includeNativeBindings = false) => Import(
            new PptxPackageSource(bytes),
            limits,
            retainImportedAssetData,
            verifiedPackageSha256,
            includeNativeBindings);

    internal static PptxImportResult Import(
        PptxPackageSource source,
        EffectiveCodecLimits limits,
        bool retainImportedAssetData = true,
        string? verifiedPackageSha256 = null,
        bool includeNativeBindings = false)
    {
        using var importStage = PpjBuildProfiler.Measure("pptx.import");
        List<PptxNativeBinding>? previewBindings = includeNativeBindings ? [] : null;
        var opaque = PackageGuards.ValidateAndCollectOpaque(
            source.OpenRead,
            source.Length,
            limits,
            OpcPackageProfile.Pptx,
            out var packagePaths);
        var packageSha256 = verifiedPackageSha256 ?? source.Sha256();
        if (source.TryGetMaterialized(out var sourceBytes))
        {
            opaque.SourcePackage = new SourcePackageSnapshot
            {
                Data = UnsafeByteOperations.UnsafeWrap(sourceBytes),
                Sha256 = packageSha256,
            };
        }
        var nativeObjects = new PptxNativeObjectCatalog(opaque, packagePaths, limits);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships for source-bound, fail-closed export from the validated package snapshot.",
                opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));

        using var stream = source.OpenRead();
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        var presentationRoot = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no presentation root.", "ppt/presentation.xml");
        var slideIds = presentationRoot.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
        if ((uint)slideIds.Length > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"PPTX presentation has {slideIds.Length} slides and exceeds max_sheets ({limits.MaxSheets}).", "ppt/presentation.xml");
        var slideParts = ResolveSlideParts(presentationPart, slideIds);
        var publicSlideIds = slideIds.Select((_, index) => $"presentation/slide/{index + 1}").ToArray();
        var publicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(slideIds
            .Select((slideId, index) => (
                RelationshipId: slideId.RelationshipId?.Value ?? string.Empty,
                PublicId: publicSlideIds[index])));
        var publicSlideIdByNativeId = BuildSectionSlideIdMap(slideIds
            .Select((slideId, index) => (
                NativeId: slideId.Id?.Value,
                PublicId: publicSlideIds[index])));
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: $"presentation/slide/{index + 1}"))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var assetCatalog = new PptxAssetCatalog(
            [],
            limits,
            nativeObjects.ValidatedPartSha256,
            retainImportedAssetData);
        var masterGraph = ReadMasterGraph(presentationPart);
        var layoutIdByPartPath = masterGraph
            .SelectMany(master => master.Layouts)
            .ToDictionary(layout => PartPath(layout.Part), layout => layout.Id, StringComparer.OrdinalIgnoreCase);
        var importedTheme = CanonicalThemePart(masterGraph)?.Theme;
        var importedAccentRgb = TryReadSourceBoundThemeAccentRgb(importedTheme);
        var importedAccent1Tint = TryReadSourceBoundThemeAccent1Tint(importedTheme);
        var importedAccent1Shade = TryReadSourceBoundThemeAccent1Shade(importedTheme);
        var importedAccent1LumMod = TryReadSourceBoundThemeAccent1LumMod(importedTheme);
        var importedAccent1LumOff = TryReadSourceBoundThemeAccent1LumOff(importedTheme);
        var importedAccent1AlphaMod = TryReadSourceBoundThemeAccent1AlphaMod(importedTheme);
        var importedAccent1AlphaOff = TryReadSourceBoundThemeAccent1AlphaOff(importedTheme);
        var importedAccent1SatMod = TryReadSourceBoundThemeAccent1SatMod(importedTheme);
        var importedAccent1SatOff = TryReadSourceBoundThemeAccent1SatOff(importedTheme);
        var importedAccent1RedMod = TryReadSourceBoundThemeAccent1RedMod(importedTheme);
        var importedAccent1RedOff = TryReadSourceBoundThemeAccent1RedOff(importedTheme);
        var importedAccent1GreenMod = TryReadSourceBoundThemeAccent1GreenMod(importedTheme);
        var importedAccent1GreenOff = TryReadSourceBoundThemeAccent1GreenOff(importedTheme);
        var importedAccent1BlueMod = TryReadSourceBoundThemeAccent1BlueMod(importedTheme);
        var importedAccent1BlueOff = TryReadSourceBoundThemeAccent1BlueOff(importedTheme);
        var importedAccent1HueMod = TryReadSourceBoundThemeAccent1HueMod(importedTheme);
        var importedAccent1HueOff = TryReadSourceBoundThemeAccent1HueOff(importedTheme);
        var importedAccent2Tint = TryReadSourceBoundThemeAccent2Tint(importedTheme);
        var importedAccent2Shade = TryReadSourceBoundThemeAccent2Shade(importedTheme);
        var importedAccent2LumMod = TryReadSourceBoundThemeAccent2LumMod(importedTheme);
        var importedAccent2LumOff = TryReadSourceBoundThemeAccent2LumOff(importedTheme);
        var importedAccent2AlphaMod = TryReadSourceBoundThemeAccent2AlphaMod(importedTheme);
        var importedAccent2AlphaOff = TryReadSourceBoundThemeAccent2AlphaOff(importedTheme);
        var importedAccent2SatMod = TryReadSourceBoundThemeAccent2SatMod(importedTheme);
        var importedAccent2RedMod = TryReadSourceBoundThemeAccent2RedMod(importedTheme);
        var importedAccent2SatOff = TryReadSourceBoundThemeAccent2SatOff(importedTheme);
        var importedAccent2RedOff = TryReadSourceBoundThemeAccent2RedOff(importedTheme);
        var importedAccent2GreenMod = TryReadSourceBoundThemeAccent2GreenMod(importedTheme);
        var importedAccent2GreenOff = TryReadSourceBoundThemeAccent2GreenOff(importedTheme);
        var importedAccent2BlueMod = TryReadSourceBoundThemeAccent2BlueMod(importedTheme);
        var importedAccent2BlueOff = TryReadSourceBoundThemeAccent2BlueOff(importedTheme);
        var importedAccent2HueMod = TryReadSourceBoundThemeAccent2HueMod(importedTheme);
        var importedAccent2HueOff = TryReadSourceBoundThemeAccent2HueOff(importedTheme);
        var importedAccent3Tint = TryReadSourceBoundThemeAccent3Tint(importedTheme);
        var importedAccent3Shade = TryReadSourceBoundThemeAccent3Shade(importedTheme);
        var importedAccent3LumMod = TryReadSourceBoundThemeAccent3LumMod(importedTheme);
        var importedAccent3LumOff = TryReadSourceBoundThemeAccent3LumOff(importedTheme);
        var importedAccent3AlphaMod = TryReadSourceBoundThemeAccent3AlphaMod(importedTheme);
        var importedAccent3AlphaOff = TryReadSourceBoundThemeAccent3AlphaOff(importedTheme);
        var importedAccent3SatMod = TryReadSourceBoundThemeAccent3SatMod(importedTheme);
        var importedAccent3SatOff = TryReadSourceBoundThemeAccent3SatOff(importedTheme);
        var importedAccent3RedMod = TryReadSourceBoundThemeAccent3RedMod(importedTheme);
        var importedAccent3RedOff = TryReadSourceBoundThemeAccent3RedOff(importedTheme);
        var importedAccent3GreenMod = TryReadSourceBoundThemeAccent3GreenMod(importedTheme);
        var importedAccent3GreenOff = TryReadSourceBoundThemeAccent3GreenOff(importedTheme);
        var importedAccent3BlueMod = TryReadSourceBoundThemeAccent3BlueMod(importedTheme);
        var importedAccent3BlueOff = TryReadSourceBoundThemeAccent3BlueOff(importedTheme);
        var importedAccent3HueMod = TryReadSourceBoundThemeAccent3HueMod(importedTheme);
        var importedAccent3HueOff = TryReadSourceBoundThemeAccent3HueOff(importedTheme);
        var importedAccent4Tint = TryReadSourceBoundThemeAccent4Tint(importedTheme);
        var importedAccent4Shade = TryReadSourceBoundThemeAccent4Shade(importedTheme);
        var importedAccent4LumMod = TryReadSourceBoundThemeAccent4LumMod(importedTheme);
        var importedAccent4LumOff = TryReadSourceBoundThemeAccent4LumOff(importedTheme);
        var importedAccent4AlphaMod = TryReadSourceBoundThemeAccent4AlphaMod(importedTheme);
        var importedAccent4AlphaOff = TryReadSourceBoundThemeAccent4AlphaOff(importedTheme);
        var importedAccent4SatMod = TryReadSourceBoundThemeAccent4SatMod(importedTheme);
        var importedAccent4SatOff = TryReadSourceBoundThemeAccent4SatOff(importedTheme);
        var importedAccent4RedMod = TryReadSourceBoundThemeAccent4RedMod(importedTheme);
        var importedAccent4GreenMod = TryReadSourceBoundThemeAccent4GreenMod(importedTheme);
        var importedAccent4BlueOff = TryReadSourceBoundThemeAccent4BlueOff(importedTheme);
        var importedAccent4GreenOff = TryReadSourceBoundThemeAccent4GreenOff(importedTheme);
        var importedAccent4BlueMod = TryReadSourceBoundThemeAccent4BlueMod(importedTheme);
        var importedAccent4HueMod = TryReadSourceBoundThemeAccent4HueMod(importedTheme);
        var importedAccent4HueOff = TryReadSourceBoundThemeAccent4HueOff(importedTheme);
        var importedAccent4RedOff = TryReadSourceBoundThemeAccent4RedOff(importedTheme);
        var importedAccent5Tint = TryReadSourceBoundThemeAccent5Tint(importedTheme);
        var importedAccent5Shade = TryReadSourceBoundThemeAccent5Shade(importedTheme);
        var importedAccent5LumMod = TryReadSourceBoundThemeAccent5LumMod(importedTheme);
        var importedAccent5LumOff = TryReadSourceBoundThemeAccent5LumOff(importedTheme);
        var importedAccent5AlphaMod = TryReadSourceBoundThemeAccent5AlphaMod(importedTheme);
        var importedAccent5AlphaOff = TryReadSourceBoundThemeAccent5AlphaOff(importedTheme);
        var importedAccent5SatMod = TryReadSourceBoundThemeAccent5SatMod(importedTheme);
        var importedAccent5SatOff = TryReadSourceBoundThemeAccent5SatOff(importedTheme);
        var importedAccent5RedMod = TryReadSourceBoundThemeAccent5RedMod(importedTheme);
        var importedAccent5RedOff = TryReadSourceBoundThemeAccent5RedOff(importedTheme);
        var importedAccent5GreenMod = TryReadSourceBoundThemeAccent5GreenMod(importedTheme);
        var importedAccent5GreenOff = TryReadSourceBoundThemeAccent5GreenOff(importedTheme);
        var importedAccent5BlueMod = TryReadSourceBoundThemeAccent5BlueMod(importedTheme);
        var importedAccent5BlueOff = TryReadSourceBoundThemeAccent5BlueOff(importedTheme);
        var importedAccent5HueMod = TryReadSourceBoundThemeAccent5HueMod(importedTheme);
        var importedAccent5HueOff = TryReadSourceBoundThemeAccent5HueOff(importedTheme);
        var importedAccent6Tint = TryReadSourceBoundThemeAccent6Tint(importedTheme);
        var importedAccent6Shade = TryReadSourceBoundThemeAccent6Shade(importedTheme);
        var importedAccent6LumMod = TryReadSourceBoundThemeAccent6LumMod(importedTheme);
        var importedColorRoleRgb = TryReadSourceBoundThemeColorRoleRgb(importedTheme);
        var importedThemeName = importedTheme?.Name?.Value;
        if (string.IsNullOrWhiteSpace(importedThemeName)) importedThemeName = null;
        var importedMajorFontFamily = importedTheme?.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMajorFontFamily)) importedMajorFontFamily = null;
        var importedMinorFontFamily = importedTheme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMinorFontFamily)) importedMinorFontFamily = null;
        var importedMajorFontFamilyEastAsia = importedTheme?.ThemeElements?.FontScheme?.MajorFont?.EastAsianFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMajorFontFamilyEastAsia)) importedMajorFontFamilyEastAsia = null;
        var importedMinorFontFamilyEastAsia = importedTheme?.ThemeElements?.FontScheme?.MinorFont?.EastAsianFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMinorFontFamilyEastAsia)) importedMinorFontFamilyEastAsia = null;
        var importedMajorFontFamilyComplexScript = importedTheme?.ThemeElements?.FontScheme?.MajorFont?.ComplexScriptFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMajorFontFamilyComplexScript)) importedMajorFontFamilyComplexScript = null;
        var importedMinorFontFamilyComplexScript = importedTheme?.ThemeElements?.FontScheme?.MinorFont?.ComplexScriptFont?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(importedMinorFontFamilyComplexScript)) importedMinorFontFamilyComplexScript = null;
        var importedThemeArtifact = new PresentationThemeArtifact();
        if (importedAccentRgb is not null) importedThemeArtifact.AccentRgb.Add(importedAccentRgb);
        if (importedAccent1Tint is { } tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                TintThousandth = tint,
            });
        if (importedAccent1Shade is { } shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                ShadeThousandth = shade,
            });
        if (importedAccent1LumMod is { } lumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                LuminanceModulationThousandth = lumMod,
            });
        if (importedAccent1LumOff is { } lumOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                LuminanceOffsetThousandth = lumOff,
            });
        if (importedAccent1AlphaMod is { } alphaMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                AlphaModulationThousandth = alphaMod,
            });
        if (importedAccent1AlphaOff is { } alphaOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                AlphaOffsetThousandth = alphaOff,
            });
        if (importedAccent1SatMod is { } satMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                SaturationModulationThousandth = satMod,
            });
        if (importedAccent1SatOff is { } satOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                SaturationOffsetThousandth = satOff,
            });
        if (importedAccent1RedMod is { } redMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                RedModulationThousandth = redMod,
            });
        if (importedAccent1RedOff is { } redOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                RedOffsetThousandth = redOff,
            });
        if (importedAccent1GreenMod is { } greenMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                GreenModulationThousandth = greenMod,
            });
        if (importedAccent1GreenOff is { } greenOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                GreenOffsetThousandth = greenOff,
            });
        if (importedAccent1BlueMod is { } blueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                BlueModulationThousandth = blueMod,
            });
        if (importedAccent1BlueOff is { } blueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                BlueOffsetThousandth = blueOff,
            });
        if (importedAccent1HueMod is { } hueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                HueModulationThousandth = hueMod,
            });
        if (importedAccent1HueOff is { } hueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent1",
                HueOffsetAngleThousandth = hueOff,
            });
        if (importedAccent2Tint is { } accent2Tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                TintThousandth = accent2Tint,
            });
        if (importedAccent2Shade is { } accent2Shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                ShadeThousandth = accent2Shade,
            });
        if (importedAccent2LumMod is { } accent2LumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                LuminanceModulationThousandth = accent2LumMod,
            });
        if (importedAccent2LumOff is { } accent2LumOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                LuminanceOffsetThousandth = accent2LumOff,
            });
        if (importedAccent2AlphaMod is { } accent2AlphaMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                AlphaModulationThousandth = accent2AlphaMod,
            });
        if (importedAccent2AlphaOff is { } accent2AlphaOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                AlphaOffsetThousandth = accent2AlphaOff,
            });
        if (importedAccent2SatMod is { } accent2SatMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                SaturationModulationThousandth = accent2SatMod,
            });
        if (importedAccent2RedMod is { } accent2RedMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                RedModulationThousandth = accent2RedMod,
            });
        if (importedAccent2SatOff is { } accent2SatOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                SaturationOffsetThousandth = accent2SatOff,
            });
        if (importedAccent2RedOff is { } accent2RedOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                RedOffsetThousandth = accent2RedOff,
            });
        if (importedAccent2GreenMod is { } accent2GreenMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                GreenModulationThousandth = accent2GreenMod,
            });
        if (importedAccent2GreenOff is { } accent2GreenOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                GreenOffsetThousandth = accent2GreenOff,
            });
        if (importedAccent2BlueMod is { } accent2BlueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                BlueModulationThousandth = accent2BlueMod,
            });
        if (importedAccent2BlueOff is { } accent2BlueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                BlueOffsetThousandth = accent2BlueOff,
            });
        if (importedAccent2HueMod is { } accent2HueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                HueModulationThousandth = accent2HueMod,
            });
        if (importedAccent2HueOff is { } accent2HueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent2",
                HueOffsetAngleThousandth = accent2HueOff,
            });
        if (importedAccent3Tint is { } accent3Tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                TintThousandth = accent3Tint,
            });
        if (importedAccent3Shade is { } accent3Shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                ShadeThousandth = accent3Shade,
            });
        if (importedAccent3LumMod is { } accent3LumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                LuminanceModulationThousandth = accent3LumMod,
            });
        if (importedAccent3LumOff is { } accent3LumOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                LuminanceOffsetThousandth = accent3LumOff,
            });
        if (importedAccent3AlphaMod is { } accent3AlphaMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                AlphaModulationThousandth = accent3AlphaMod,
            });
        if (importedAccent3AlphaOff is { } accent3AlphaOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                AlphaOffsetThousandth = accent3AlphaOff,
            });
        if (importedAccent3SatMod is { } accent3SatMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                SaturationModulationThousandth = accent3SatMod,
            });
        if (importedAccent3SatOff is { } accent3SatOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                SaturationOffsetThousandth = accent3SatOff,
            });
        if (importedAccent3RedMod is { } accent3RedMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                RedModulationThousandth = accent3RedMod,
            });
        if (importedAccent3RedOff is { } accent3RedOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                RedOffsetThousandth = accent3RedOff,
            });
        if (importedAccent3GreenMod is { } accent3GreenMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                GreenModulationThousandth = accent3GreenMod,
            });
        if (importedAccent3GreenOff is { } accent3GreenOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                GreenOffsetThousandth = accent3GreenOff,
            });
        if (importedAccent3BlueMod is { } accent3BlueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                BlueModulationThousandth = accent3BlueMod,
            });
        if (importedAccent3BlueOff is { } accent3BlueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                BlueOffsetThousandth = accent3BlueOff,
            });
        if (importedAccent3HueMod is { } accent3HueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                HueModulationThousandth = accent3HueMod,
            });
        if (importedAccent3HueOff is { } accent3HueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent3",
                HueOffsetAngleThousandth = accent3HueOff,
            });
        if (importedAccent4HueOff is { } accent4HueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                HueOffsetAngleThousandth = accent4HueOff,
            });
        if (importedAccent4Tint is { } accent4Tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                TintThousandth = accent4Tint,
            });
        if (importedAccent4Shade is { } accent4Shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                ShadeThousandth = accent4Shade,
            });
        if (importedAccent4LumMod is { } accent4LumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                LuminanceModulationThousandth = accent4LumMod,
            });
        if (importedAccent4LumOff is { } accent4LumOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                LuminanceOffsetThousandth = accent4LumOff,
            });
        if (importedAccent4AlphaMod is { } accent4AlphaMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                AlphaModulationThousandth = accent4AlphaMod,
            });
        if (importedAccent4AlphaOff is { } accent4AlphaOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                AlphaOffsetThousandth = accent4AlphaOff,
            });
        if (importedAccent4SatMod is { } accent4SatMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                SaturationModulationThousandth = accent4SatMod,
            });
        if (importedAccent4SatOff is { } accent4SatOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                SaturationOffsetThousandth = accent4SatOff,
            });
        if (importedAccent4RedMod is { } accent4RedMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                RedModulationThousandth = accent4RedMod,
            });
        if (importedAccent4GreenMod is { } accent4GreenMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                GreenModulationThousandth = accent4GreenMod,
            });
        if (importedAccent4BlueMod is { } accent4BlueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                BlueModulationThousandth = accent4BlueMod,
            });
        if (importedAccent4HueMod is { } accent4HueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                HueModulationThousandth = accent4HueMod,
            });
        if (importedAccent4BlueOff is { } accent4BlueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                BlueOffsetThousandth = accent4BlueOff,
            });
        if (importedAccent4GreenOff is { } accent4GreenOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                GreenOffsetThousandth = accent4GreenOff,
            });
        if (importedAccent4RedOff is { } accent4RedOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent4",
                RedOffsetThousandth = accent4RedOff,
            });
        if (importedAccent5Tint is { } accent5Tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                TintThousandth = accent5Tint,
            });
        if (importedAccent5Shade is { } accent5Shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                ShadeThousandth = accent5Shade,
            });
        if (importedAccent5LumMod is { } accent5LumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                LuminanceModulationThousandth = accent5LumMod,
            });
        if (importedAccent5LumOff is { } accent5LumOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                LuminanceOffsetThousandth = accent5LumOff,
            });
        if (importedAccent5AlphaMod is { } accent5AlphaMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                AlphaModulationThousandth = accent5AlphaMod,
            });
        if (importedAccent5AlphaOff is { } accent5AlphaOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                AlphaOffsetThousandth = accent5AlphaOff,
            });
        if (importedAccent5SatMod is { } accent5SatMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                SaturationModulationThousandth = accent5SatMod,
            });
        if (importedAccent5SatOff is { } accent5SatOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                SaturationOffsetThousandth = accent5SatOff,
            });
        if (importedAccent5RedMod is { } accent5RedMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                RedModulationThousandth = accent5RedMod,
            });
        if (importedAccent5RedOff is { } accent5RedOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                RedOffsetThousandth = accent5RedOff,
            });
        if (importedAccent5GreenMod is { } accent5GreenMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                GreenModulationThousandth = accent5GreenMod,
            });
        if (importedAccent5GreenOff is { } accent5GreenOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                GreenOffsetThousandth = accent5GreenOff,
            });
        if (importedAccent5BlueMod is { } accent5BlueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                BlueModulationThousandth = accent5BlueMod,
            });
        if (importedAccent5BlueOff is { } accent5BlueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                BlueOffsetThousandth = accent5BlueOff,
            });
        if (importedAccent5HueMod is { } accent5HueMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                HueModulationThousandth = accent5HueMod,
            });
        if (importedAccent5HueOff is { } accent5HueOff)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent5",
                HueOffsetAngleThousandth = accent5HueOff,
            });
        if (importedAccent6Tint is { } accent6Tint)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent6",
                TintThousandth = accent6Tint,
            });
        if (importedAccent6Shade is { } accent6Shade)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent6",
                ShadeThousandth = accent6Shade,
            });
        if (importedAccent6LumMod is { } accent6LumMod)
            importedThemeArtifact.AccentTransforms.Add(new PresentationThemeColorTransform
            {
                Role = "accent6",
                LuminanceModulationThousandth = accent6LumMod,
            });
        if (importedColorRoleRgb is not null)
        {
            importedThemeArtifact.Dark1Rgb = importedColorRoleRgb[0];
            importedThemeArtifact.Light1Rgb = importedColorRoleRgb[1];
            importedThemeArtifact.Dark2Rgb = importedColorRoleRgb[2];
            importedThemeArtifact.Light2Rgb = importedColorRoleRgb[3];
            importedThemeArtifact.HyperlinkRgb = importedColorRoleRgb[4];
            importedThemeArtifact.FollowedHyperlinkRgb = importedColorRoleRgb[5];
        }
        if (importedThemeName is not null) importedThemeArtifact.Name = importedThemeName;
        if (importedMajorFontFamily is not null) importedThemeArtifact.MajorFontFamily = importedMajorFontFamily;
        if (importedMinorFontFamily is not null) importedThemeArtifact.MinorFontFamily = importedMinorFontFamily;
        if (importedMajorFontFamilyEastAsia is not null) importedThemeArtifact.MajorFontFamilyEastAsia = importedMajorFontFamilyEastAsia;
        if (importedMinorFontFamilyEastAsia is not null) importedThemeArtifact.MinorFontFamilyEastAsia = importedMinorFontFamilyEastAsia;
        if (importedMajorFontFamilyComplexScript is not null) importedThemeArtifact.MajorFontFamilyComplexScript = importedMajorFontFamilyComplexScript;
        if (importedMinorFontFamilyComplexScript is not null) importedThemeArtifact.MinorFontFamilyComplexScript = importedMinorFontFamilyComplexScript;

        var artifact = new PresentationArtifact
        {
            Id = "presentation/1",
            Name = "Imported presentation",
            SlideWidthEmu = presentationRoot.SlideSize?.Cx?.Value ?? DefaultSlideWidthEmu,
            SlideHeightEmu = presentationRoot.SlideSize?.Cy?.Value ?? DefaultSlideHeightEmu,
            AuthoredTheme = importedThemeArtifact.AccentRgb.Count == 0 &&
                importedThemeArtifact.AccentTransforms.Count == 0 &&
                !importedThemeArtifact.HasDark1Rgb &&
                !importedThemeArtifact.HasLight1Rgb &&
                !importedThemeArtifact.HasDark2Rgb &&
                !importedThemeArtifact.HasLight2Rgb &&
                !importedThemeArtifact.HasHyperlinkRgb &&
                !importedThemeArtifact.HasFollowedHyperlinkRgb &&
                !importedThemeArtifact.HasName &&
                !importedThemeArtifact.HasMajorFontFamily &&
                !importedThemeArtifact.HasMinorFontFamily &&
                !importedThemeArtifact.HasMajorFontFamilyEastAsia &&
                !importedThemeArtifact.HasMinorFontFamilyEastAsia &&
                !importedThemeArtifact.HasMajorFontFamilyComplexScript &&
                !importedThemeArtifact.HasMinorFontFamilyComplexScript
                ? null
                : importedThemeArtifact,
        };
        artifact.ViewProperties = PptxViewPropertiesCodec.Read(presentationPart);
        ulong semanticItems = 0;
        var customShows = PptxCustomShowCodec.Read(presentationPart, publicSlideIdByRelationshipId, limits);
        semanticItems = checked(semanticItems + customShows.SemanticItems);
        if (semanticItems > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", "ppt/presentation.xml");
        artifact.CustomShows.Add(customShows.Shows);
        artifact.CustomShowsOpaque = customShows.Opaque;
        if (customShows.Opaque)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_presentation_custom_shows_retained",
                $"Retained an unsupported custom-show graph without exposing incomplete editable semantics: {customShows.Reason}.",
                "ppt/presentation.xml"));
        var sections = PptxSectionCodec.Read(presentationPart, publicSlideIdByNativeId, publicSlideIds, limits);
        semanticItems = checked(semanticItems + sections.SemanticItems);
        if (semanticItems > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", "ppt/presentation.xml");
        artifact.Sections.Add(sections.Sections);
        artifact.SectionsOpaque = sections.Opaque;
        if (sections.Opaque)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_presentation_sections_retained",
                $"Retained an unsupported PowerPoint section graph without exposing incomplete editable semantics: {sections.Reason}.",
                "ppt/presentation.xml"));
        var customShowCatalog = PptxCustomShowCatalog.From(customShows.Shows);
        foreach (var master in masterGraph)
        {
            var masterRoot = master.Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"Presentation master {master.Index + 1} has no slide master root.", PartPath(master.Part));
            var masterCommon = masterRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation master {master.Index + 1} has no common slide data.", PartPath(master.Part));
            var masterShapeTree = masterCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation master {master.Index + 1} has no shape tree.", PartPath(master.Part));
            var masterContext = new PptxPartContext(master.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
            var textStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
            var background = PptxBackgroundCodec.Read(masterCommon, masterContext);
            var masterArtifact = new PresentationMaster
            {
                Id = master.Id,
                Name = masterCommon.Name?.Value ?? $"Master {master.Index + 1}",
                TextStyles = textStyles,
                Source = new PresentationMasterSourceBinding
                {
                    MasterIndex = checked((uint)master.Index),
                    PartPath = PartPath(master.Part),
                    RelationshipId = master.RelationshipId,
                    MasterXmlSha256 = HashElement(masterRoot),
                    TextStylesSemanticSha256 = MasterTextStylesSemanticHash(textStyles),
                    TextStylesEditable = PptxMasterTextStylesCodec.Supports(masterRoot),
                    BackgroundSemanticSha256 = BackgroundSemanticHash(background),
                    BackgroundEditable = PptxBackgroundCodec.Supports(masterCommon, masterContext),
                },
            };
            if (background is not null) masterArtifact.Background = background;
            masterArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(masterShapeTree, master.Id, masterContext));
            artifact.Masters.Add(masterArtifact);
            foreach (var layout in master.Layouts)
            {
                var layoutRoot = layout.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no slide layout root.", PartPath(layout.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no common slide data.", PartPath(layout.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layout.Index + 1} under master {master.Index + 1} has no shape tree.", PartPath(layout.Part));
                var layoutContext = new PptxPartContext(layout.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
                var layoutBackground = PptxBackgroundCodec.Read(layoutCommon, layoutContext);
                var layoutArtifact = new PresentationLayout
                {
                    Id = layout.Id,
                    Name = layoutCommon.Name?.Value ?? $"Layout {layout.Index + 1}",
                    MasterId = master.Id,
                    Type = LayoutTypeName(layoutRoot),
                    Source = new PresentationLayoutSourceBinding
                    {
                        LayoutIndex = checked((uint)layout.Index),
                        PartPath = PartPath(layout.Part),
                        RelationshipId = layout.RelationshipId,
                        LayoutXmlSha256 = HashElement(layoutRoot),
                        BackgroundSemanticSha256 = BackgroundSemanticHash(layoutBackground),
                        BackgroundEditable = PptxBackgroundCodec.Supports(layoutCommon, layoutContext),
                    },
                };
                if (layoutBackground is not null) layoutArtifact.Background = layoutBackground;
                layoutArtifact.Placeholders.Add(PptxPlaceholderCodec.Read(layoutShapeTree, layout.Id, layoutContext));
                artifact.Layouts.Add(layoutArtifact);
                layout.Part.UnloadRootElement();
            }
            master.Part.UnloadRootElement();
        }
        for (var slideIndex = 0; slideIndex < slideIds.Length; slideIndex++)
        {
            var slideId = slideIds[slideIndex];
            var relationshipId = slideId.RelationshipId?.Value ?? string.Empty;
            var slidePart = slideParts[slideIndex];
            var slideRoot = slidePart.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
            var slideCommon = slideRoot.CommonSlideData ??
                throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
            var shapeTree = slideCommon.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
            var slideBackground = PptxBackgroundCodec.Read(slideCommon, slideContext);
            var slideTransition = PptxTransitionCodec.Read(slideRoot);
            var slideVisibility = PptxSlideVisibilityCodec.Read(slideRoot);
            var elements = ShapeElements(shapeTree);
            var deletionAnalysis = PptxElementDeletionCodec.AnalyzeSlide(slidePart);
            var zOrderPlan = AnalyzeElementZOrder(elements);
            var slideArtifactId = $"presentation/slide/{slideIndex + 1}";
            var elementIdsByNativeId = NativeElementIds(elements, slideArtifactId, previewBindings, PartPath(slidePart));
            P.Slide? previousSlideRoot = null;
            IReadOnlyDictionary<uint, string>? previousElementIdsByNativeId = null;
            string? previousSlideArtifactId = null;
            if (slideIndex > 0)
            {
                previousSlideRoot = slideParts[slideIndex - 1].Slide;
                var previousShapeTree = previousSlideRoot?.CommonSlideData?.ShapeTree;
                previousSlideArtifactId = $"presentation/slide/{slideIndex}";
                if (previousShapeTree is not null)
                    previousElementIdsByNativeId = NativeElementIds(ShapeElements(previousShapeTree), previousSlideArtifactId);
            }
            var slideTiming = PptxTimingCodec.Read(
                slideRoot,
                elementIdsByNativeId,
                previousSlideRoot,
                previousElementIdsByNativeId,
                previousSlideArtifactId);
            var sourceEntry = new PptxSourceSlideEntry(slideIndex, slideId, relationshipId, slidePart);
            var deletionPlan = PptxSlideDeletionCodec.Analyze(presentationPart, sourceEntry, opaque);
            var clonePlan = PptxSlideCloneCodec.Analyze(presentationPart, sourceEntry, slideParts.ToHashSet());
            var target = new PresentationSlide
            {
                Id = slideArtifactId,
                // An absent p:cSld/@name is a real source value. Do not invent
                // a display name here: source-preserving slide clone and edit
                // checks must be able to distinguish absent from authored text.
                Name = slideRoot.CommonSlideData?.Name?.Value ?? string.Empty,
                LayoutId = slidePart.SlideLayoutPart is { } layoutPart
                    ? layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart)) ??
                      throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart))
                    : string.Empty,
                Source = new PresentationSlideSourceBinding
                {
                    SlideIndex = checked((uint)slideIndex),
                    PartPath = PartPath(slidePart),
                    RelationshipId = relationshipId,
                    SlideXmlSha256 = HashElement(slideRoot),
                    LayoutRelationshipId = slidePart.SlideLayoutPart is { } boundLayout ? slidePart.GetIdOfPart(boundLayout) : string.Empty,
                    BackgroundSemanticSha256 = BackgroundSemanticHash(slideBackground),
                    BackgroundEditable = PptxBackgroundCodec.Supports(slideCommon, slideContext),
                    SpeakerNotesAddable = PptxSpeakerNotesCodec.CanAddSourceBound(presentationPart, slidePart),
                    LegacyCommentsAddable = PptxLegacyCommentsCodec.CanAddSourceBound(presentationPart, slidePart),
                    LegacyCommentsEditable = PptxLegacyCommentsCodec.CanEditSourceBound(presentationPart, slidePart, slideIndex),
                    CommentPartPresent = PptxLegacyCommentsCodec.CommentPartPresent(slidePart),
                    CommentFamily = PptxLegacyCommentsCodec.CommentFamily(presentationPart),
                    TransitionSemanticSha256 = PptxTransitionCodec.SemanticHash(slideTransition),
                    TransitionEditable = PptxTransitionCodec.Supports(slideRoot),
                    TransitionPresent = PptxTransitionCodec.HasTransition(slideRoot),
                    TransitionAddable = PptxTransitionCodec.CanAdd(slideRoot),
                    TimingPresent = slideTiming.Present,
                    TimingEditable = slideTiming.Editable,
                    TimingAddable = slideTiming.Addable,
                    TimingSemanticSha256 = slideTiming.SemanticSha256,
                    VisibilitySemanticSha256 = slideVisibility.SemanticSha256,
                    VisibilityEditable = slideVisibility.Editable,
                    DeletionCapability = new PresentationSlideDeletionCapability
                    {
                        Supported = deletionPlan.Supported,
                        BlockedReason = deletionPlan.BlockedReason,
                        OwnedPartCount = deletionPlan.OwnedPartCount,
                    },
                    CloneCapability = new PresentationSlideCloneCapability
                    {
                        Supported = clonePlan.Supported,
                        BlockedReason = clonePlan.BlockedReason,
                        ClonedPartCount = clonePlan.ClonedPartCount,
                        SharedPartCount = clonePlan.SharedPartCount,
                    },
                },
            };
            if (slideVisibility.Hidden is { } hidden) target.Hidden = hidden;
            if (slideBackground is not null) target.Background = slideBackground;
            if (slideTransition is not null) target.Transition = slideTransition;
            target.Animations.AddRange(slideTiming.Animations);
            if (slideTiming.Morph is not null) target.Morph = slideTiming.Morph;
            if (PptxSpeakerNotesCodec.Read(slidePart) is { } speakerNotes)
                target.SpeakerNotes = speakerNotes;
            target.LegacyComments.Add(PptxLegacyCommentsCodec.Read(presentationPart, slidePart, slideIndex, diagnostics));
            target.ModernComments.Add(PptxModernCommentsCodec.Read(
                presentationPart,
                slideId,
                slidePart,
                elements,
                elementIdsByNativeId,
                slideIndex,
                diagnostics));
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                semanticItems++;
                if (semanticItems > limits.MaxCells)
                    throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                // Third-party decks commonly place line separators just
                // outside the canvas (for example x=-591436). ReadElement
                // applies that exception only to line geometry; ordinary
                // imported shapes remain subject to the normal frame rule.
                var importedElement = ReadElement(elements[elementIndex], slideIndex, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                if (importedElement.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    semanticItems += checked((ulong)importedElement.Table.Rows.Sum(row => row.Cells.Count));
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                }
                SetElementDeletionCapability(
                    importedElement,
                    PptxElementDeletionCodec.Analyze(slidePart, elements[elementIndex], elements, deletionAnalysis));
                SetElementZOrderCapability(importedElement, zOrderPlan);
                target.Elements.Add(importedElement);
            }
            // PPJ exposes the direct shape-tree order explicitly as the
            // page's initial reading-order permutation.  It is intentionally
            // copied from the native order rather than inferred later from
            // projected z-order or nested component expansion.
            target.ReadingOrder.Add(target.Elements.Select(element => element.Id));
            artifact.Slides.Add(target);
            if (slideIndex > 0)
                slideParts[slideIndex - 1].UnloadRootElement();
        }
        if (slideParts.Length > 0)
            slideParts[^1].UnloadRootElement();
        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecWireProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Presentation = artifact,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "pptx",
                PackageSha256 = packageSha256,
                Producer = "office-kit/OfficeKit",
            },
        };
        envelope.Assets.Add(assetCatalog.ImportedAssets);
        envelope.Diagnostics.Add(diagnostics);
        return new PptxImportResult(envelope, diagnostics) { NativeBindings = previewBindings ?? [] };
    }

    internal static void HydrateSourceImageAssets(ArtifactEnvelope envelope, PptxPackageSource source)
    {
        if (!envelope.Assets.Any(asset => asset.Data.IsEmpty)) return;
        using var stream = source.OpenRead();
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var presentationPart = package.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        PptxAssetCatalog.HydrateSourceImages(DescendantImageParts(presentationPart), envelope.Assets);
    }

    private static IEnumerable<ImagePart> DescendantImageParts(OpenXmlPart root)
    {
        var pending = new Stack<OpenXmlPart>(root.Parts.Select(pair => pair.OpenXmlPart));
        var seen = new HashSet<OpenXmlPart>();
        while (pending.TryPop(out var part))
        {
            if (!seen.Add(part)) continue;
            if (part is ImagePart image) yield return image;
            foreach (var child in part.Parts)
                pending.Push(child.OpenXmlPart);
        }
    }

    internal static PptxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        var requiresSourcePreservation =
            envelope.ProtocolVersion == CodecWireProtocol.ProtocolVersion &&
            envelope.Family == ArtifactFamily.Presentation &&
            envelope.PayloadCase == ArtifactEnvelope.PayloadOneofCase.Presentation &&
            RequiresSourcePreservation(envelope);
        if (requiresSourcePreservation && envelope.OpaqueOpc?.SourcePackage is not { Data.IsEmpty: false })
            throw new CodecException(
                "missing_source_package",
                "Source-bound PPTX export requires its validated source package snapshot.");

        var assetCatalog = ValidateEnvelope(
            envelope,
            limits,
            allowEmptyExistingPictureAssets: requiresSourcePreservation);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (requiresSourcePreservation)
            return ExportPreservingSource(envelope, limits, opaqueCount, assetCatalog);

        var diagnostics = new List<Diagnostic>();

        using var stream = new MemoryStream();
        using (var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, autoSave: true))
            BuildPresentation(package, new PptxArtifactSourceFreeBuildPlan(envelope.Presentation), assetCatalog);
        var bytes = NormalizeSourceFreePackage(stream.ToArray());
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new PptxExportResult(bytes, diagnostics);
    }

    internal static PptxExportResult ExportSourceFree(
        IPptxSourceFreeBuildPlan plan,
        IReadOnlyList<Asset> assets,
        EffectiveCodecLimits limits,
        Action<Dictionary<string, byte[]>> enrichPackage)
    {
        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecWireProtocol.ProtocolVersion,
            Family = ArtifactFamily.Presentation,
            Presentation = plan.Presentation,
        };
        envelope.Assets.Add(assets);
        var validation = ValidateEnvelopeHeader(envelope, limits);

        using var stream = new MemoryStream();
        using (var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, autoSave: true))
            BuildPresentation(
                package,
                plan,
                validation.AssetCatalog,
                (slideIndex, slide) => ValidatePresentationSlide(slide, slideIndex, validation, limits));
        var bytes = NormalizeSourceFreePackage(stream.ToArray(), enrichPackage);
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new PptxExportResult(bytes, []);
    }

    private static byte[] NormalizeSourceFreePackage(
        byte[] bytes,
        Action<Dictionary<string, byte[]>>? enrichPackage = null)
    {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using (var source = new MemoryStream(bytes, writable: false))
        using (var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: false))
        {
            foreach (var entry in archive.Entries)
            {
                using var entryStream = entry.Open();
                using var copy = new MemoryStream();
                entryStream.CopyTo(copy);
                if (!parts.TryAdd(entry.FullName, copy.ToArray()))
                    throw new CodecException("duplicate_package_part", $"Generated PPTX contains duplicate part {entry.FullName}.", entry.FullName);
            }
        }

        var rootRelationshipsPath = parts.Keys.SingleOrDefault(path =>
            path.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase));
        if (rootRelationshipsPath is null)
            throw new CodecException("missing_package_relationships", "Generated PPTX has no package relationship part.", "_rels/.rels");
        XDocument rootRelationships;
        using (var rootRelationshipsStream = new MemoryStream(parts[rootRelationshipsPath], writable: false))
            rootRelationships = XDocument.Load(rootRelationshipsStream, LoadOptions.PreserveWhitespace);
        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        var officeDocumentRelationship = rootRelationships.Root?
            .Elements(relationshipsNamespace + "Relationship")
            .SingleOrDefault(relationship =>
                relationship.Attribute("Type")?.Value.EndsWith("/officeDocument", StringComparison.Ordinal) == true) ??
            throw new CodecException("missing_presentation_relationship", "Generated PPTX has no officeDocument package relationship.", "_rels/.rels");
        var reservedIds = rootRelationships.Root!
            .Elements(relationshipsNamespace + "Relationship")
            .Where(relationship => !ReferenceEquals(relationship, officeDocumentRelationship))
            .Select(relationship => relationship.Attribute("Id")?.Value ?? string.Empty)
            .ToHashSet(StringComparer.Ordinal);
        var relationshipId = "rIdOfficeKitPresentation";
        for (var suffix = 1; reservedIds.Contains(relationshipId); suffix++)
            relationshipId = $"rIdOfficeKitPresentation{suffix}";
        officeDocumentRelationship.SetAttributeValue("Id", relationshipId);
        using (var normalized = new MemoryStream())
        using (var writer = XmlWriter.Create(normalized, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false,
        }))
        {
            rootRelationships.Save(writer);
            writer.Flush();
            parts[rootRelationshipsPath] = normalized.ToArray();
        }

        enrichPackage?.Invoke(parts);

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var timestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            foreach (var part in parts.OrderBy(part => part.Key, StringComparer.Ordinal))
            {
                var entry = archive.CreateEntry(part.Key, CompressionLevel.Optimal);
                entry.LastWriteTime = timestamp;
                using var target = entry.Open();
                target.Write(part.Value);
            }
        }
        return output.ToArray();
    }

    private static bool RequiresSourcePreservation(ArtifactEnvelope envelope)
    {
        if (envelope.Source is not null) return true;
        if (envelope.OpaqueOpc is { } opaque &&
            (opaque.SourcePackage is not null || opaque.Parts.Count > 0 || opaque.PackageRelationships.Count > 0))
            return true;

        var presentation = envelope.Presentation;
        return presentation.Masters.Any(master =>
                   master.Source is not null || master.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Layouts.Any(layout =>
                   layout.Source is not null || layout.Placeholders.Any(placeholder => placeholder.Source is not null)) ||
               presentation.Slides.Any(slide =>
                   slide.Source is not null || slide.Elements.Any(element =>
                       element.Source is not null || element.ContentCase == PresentationElement.ContentOneofCase.Opaque)) ||
               presentation.ViewProperties?.Source is not null ||
               presentation.CustomShowsOpaque ||
               presentation.CustomShows.Any(show => show.Source is not null) ||
               presentation.SectionsOpaque ||
               presentation.Sections.Any(section => section.Source is not null);
    }

    private static PptxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount, PptxAssetCatalog assetCatalog)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(envelope.OpaqueOpc, envelope.Source, limits, OpcPackageProfile.Pptx);
        var nativeObjects = new PptxNativeObjectCatalog(envelope.OpaqueOpc, sourceBytes, limits);
        var growthAllowance = Math.Min(4 * 1024 * 1024, Math.Max(64 * 1024, sourceBytes.Length / 8));
        using var stream = new MemoryStream(checked(sourceBytes.Length + growthAllowance));
        stream.Write(sourceBytes);
        stream.Position = 0;
        var changedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var addedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var addedPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var replacedOpaquePartHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var removedSourcePartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var removedSourceRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        var removedElementPartPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var removedElementRelationshipKeys = new HashSet<string>(StringComparer.Ordinal);
        var clonedPartSourcePaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var clonedPackageEntryPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var authoredOverlayXmlByPartPath = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        using (var package = PresentationDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var presentationPart = package.PresentationPart ??
                throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
            var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
            var targetSlides = BindSourcePreservingSlides(presentationPart, slideIds, envelope.Presentation.Slides);
            AssertCloneOriginsRetained(targetSlides);
            var retainedTargets = targetSlides.Where(target => !target.IsClone).ToArray();
            var slideParts = retainedTargets.Select(target => target.Source.Part).ToArray();
            var slideIdByPartPath = retainedTargets
                .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
                .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
            var slidePartById = retainedTargets
                .Select(target => (Part: target.Source.Part, Id: target.Target.Id))
                .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
            var customShowCatalog = PptxCustomShowCatalog.From(envelope.Presentation.CustomShows);
            var retainedPublicSlideIdByRelationshipId = retainedTargets
                .Select(target => (target.Source.RelationshipId, target.Target.Id))
                .ToDictionary(item => item.RelationshipId, item => item.Id, StringComparer.Ordinal);
            var sourcePublicSlideIds = slideIds
                .Select(slideId => retainedPublicSlideIdByRelationshipId.GetValueOrDefault(slideId.RelationshipId?.Value ?? string.Empty) ?? string.Empty)
                .ToArray();
            var sourcePublicSlideIdByNativeId = BuildSectionSlideIdMap(slideIds
                .Select(slideId => (
                    NativeId: slideId.Id?.Value,
                    PublicId: retainedPublicSlideIdByRelationshipId.GetValueOrDefault(slideId.RelationshipId?.Value ?? string.Empty) ?? string.Empty)));
            if (targetSlides.Any(target => target.IsClone))
            {
                var sourcePublicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(retainedTargets.Select(target => (
                    RelationshipId: target.Source.RelationshipId,
                    PublicId: target.Target.Id)));
                PptxCustomShowCodec.AssertMembershipUnchangedForSlideClone(
                    presentationPart,
                    envelope.Presentation,
                    sourcePublicSlideIdByRelationshipId,
                    limits);
                PptxSectionCodec.AssertNoSectionCloneCombination(
                    presentationPart,
                    envelope.Presentation,
                    sourcePublicSlideIdByNativeId,
                    sourcePublicSlideIds,
                    limits);
            }
            var masterGraph = ReadMasterGraph(presentationPart);
            if (masterGraph.Length != envelope.Presentation.Masters.Count)
                throw new CodecException(
                    "presentation_master_topology_changed",
                    $"Source-preserving PPTX export requires the original {masterGraph.Length}-master topology; the artifact contains {envelope.Presentation.Masters.Count} masters.",
                    "ppt/presentation.xml");
            var layoutGraph = masterGraph.SelectMany(master => master.Layouts.Select(layout => (Master: master, Layout: layout))).ToArray();
            if (layoutGraph.Length != envelope.Presentation.Layouts.Count)
                throw new CodecException(
                    "presentation_layout_topology_changed",
                    $"Source-preserving PPTX export requires the original {layoutGraph.Length}-layout topology; the artifact contains {envelope.Presentation.Layouts.Count} layouts.",
                    "ppt/presentation.xml");
            ApplySourceBoundThemeFields(envelope.Presentation.AuthoredTheme, masterGraph, changedParts, replacedOpaquePartHashes);
            var layoutIdByPartPath = layoutGraph.ToDictionary(item => PartPath(item.Layout.Part), item => item.Layout.Id, StringComparer.OrdinalIgnoreCase);
            if (PptxViewPropertiesCodec.ApplySourceBound(presentationPart, envelope.Presentation.ViewProperties) is { } viewPropertiesChange)
            {
                changedParts.Add(viewPropertiesChange.PartPath);
                replacedOpaquePartHashes.Add(viewPropertiesChange.PartPath, viewPropertiesChange.Sha256);
            }
            PptxLegacyCommentsCodec.AssertSourceUnchanged(presentationPart, slideParts, retainedTargets.Select(target => target.Target).ToArray());
            CloneRequestedSourceSlides(
                presentationPart,
                targetSlides,
                layoutIdByPartPath,
                slideIdByPartPath,
                assetCatalog,
                customShowCatalog,
                nativeObjects,
                changedParts,
                addedRelationshipIds,
                addedPartPaths,
                clonedPackageEntryPaths,
                clonedPartSourcePaths);
            DeleteUnrequestedSourceSlides(
                presentationPart,
                slideIds,
                targetSlides,
                envelope.OpaqueOpc,
                changedParts,
                removedSourcePartPaths);
            var sourceRelationshipIds = slideIds.Select(slideId => slideId.RelationshipId?.Value ?? string.Empty).ToArray();
            var targetRelationshipIds = targetSlides.Select(target => target.OutputSlideId.RelationshipId?.Value ?? string.Empty).ToArray();
            var pureSourceReorder = targetSlides.Length == slideIds.Length &&
                                    targetSlides.All(target => !target.IsClone) &&
                                    sourceRelationshipIds.ToHashSet(StringComparer.Ordinal).SetEquals(targetRelationshipIds) &&
                                    !sourceRelationshipIds.SequenceEqual(targetRelationshipIds, StringComparer.Ordinal);
            var sectionsAppliedBeforeReorder = false;
            if (pureSourceReorder)
            {
                if (PptxSectionCodec.ApplySourceBound(
                        presentationPart,
                        envelope.Presentation,
                        sourcePublicSlideIdByNativeId,
                        sourcePublicSlideIds,
                        envelope.Presentation.Slides.Select(slide => slide.Id).ToArray(),
                        limits))
                    changedParts.Add(PartPath(presentationPart));
                sectionsAppliedBeforeReorder = true;
            }
            if (ReorderSourceSlideIdList(presentationPart, targetSlides))
                changedParts.Add(PartPath(presentationPart));
            if (ApplySourceBoundSlideSize(presentationPart, envelope.Presentation))
                changedParts.Add(PartPath(presentationPart));
            if (presentationPart.Presentation?.CustomShowList is not null ||
                envelope.Presentation.CustomShowsOpaque ||
                envelope.Presentation.CustomShows.Count > 0)
            {
                var outputSlideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
                if (outputSlideIds.Length != envelope.Presentation.Slides.Count)
                    throw new CodecException(
                        "presentation_slide_topology_changed",
                        "Source-preserving PPTX export could not bind custom shows to the requested slide topology.",
                        "ppt/presentation.xml");
                var publicSlideIdByRelationshipId = BuildCustomShowSlideIdMap(outputSlideIds
                    .Select((slideId, index) => (
                        RelationshipId: slideId.RelationshipId?.Value ?? string.Empty,
                        PublicId: envelope.Presentation.Slides[index].Id)));
                if (PptxCustomShowCodec.ApplySourceBound(presentationPart, envelope.Presentation, publicSlideIdByRelationshipId, limits))
                    changedParts.Add(PartPath(presentationPart));
            }
            if (!sectionsAppliedBeforeReorder)
            {
                var outputSectionSlideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
                if (outputSectionSlideIds.Length != envelope.Presentation.Slides.Count)
                    throw new CodecException(
                        "presentation_slide_topology_changed",
                        "Source-preserving PPTX export could not bind PowerPoint sections to the requested slide topology.",
                        "ppt/presentation.xml");
                var outputPublicSlideIdByNativeId = BuildSectionSlideIdMap(outputSectionSlideIds
                    .Select((slideId, index) => (
                        NativeId: slideId.Id?.Value,
                        PublicId: envelope.Presentation.Slides[index].Id)));
                var outputSlideIds = envelope.Presentation.Slides.Select(slide => slide.Id).ToArray();
                if (PptxSectionCodec.ApplySourceBound(
                        presentationPart,
                        envelope.Presentation,
                        outputPublicSlideIdByNativeId,
                        outputSlideIds,
                        outputSlideIds,
                        limits))
                    changedParts.Add(PartPath(presentationPart));
            }
            assetCatalog.IndexExistingParts(DescendantImageParts(presentationPart));

            for (var masterIndex = 0; masterIndex < masterGraph.Length; masterIndex++)
            {
                var graph = masterGraph[masterIndex];
                var masterRoot = graph.Part.SlideMaster ??
                    throw new CodecException("missing_slide_master_root", $"Presentation master {masterIndex + 1} has no slide master root.", PartPath(graph.Part));
                var masterCommon = masterRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation master {masterIndex + 1} has no common slide data.", PartPath(graph.Part));
                var masterShapeTree = masterCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation master {masterIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Masters[masterIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_master_binding",
                    $"Presentation master {masterIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                if (target.Id != graph.Id ||
                    binding.MasterIndex != masterIndex ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.MasterXmlSha256.Equals(HashElement(masterRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_binding_mismatch",
                        $"Presentation master {masterIndex + 1} does not match its hash-bound source master.",
                        PartPath(graph.Part));
                var sourceName = masterCommon.Name?.Value ?? $"Master {masterIndex + 1}";
                if (!target.Name.Equals(sourceName, StringComparison.Ordinal))
                    throw new CodecException("unsupported_presentation_edit", $"Source-preserving PPTX export cannot rename master {masterIndex + 1}.", PartPath(graph.Part));
                var masterContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
                var originalStyles = PptxMasterTextStylesCodec.Read(masterRoot, masterContext);
                var originalSemanticHash = MasterTextStylesSemanticHash(originalStyles);
                if (!binding.TextStylesSemanticSha256.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_semantics_mismatch",
                        $"Presentation master {masterIndex + 1} text styles do not match their source binding.",
                        PartPath(graph.Part));
                var requestedSemanticHash = MasterTextStylesSemanticHash(target.TextStyles);
                if (!requestedSemanticHash.Equals(originalSemanticHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.TextStylesEditable || !PptxMasterTextStylesCodec.Supports(masterRoot))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} text styles are preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxMasterTextStylesCodec.Apply(masterRoot, target.TextStyles ?? new PresentationMasterTextStyles(), masterContext);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                var originalBackground = PptxBackgroundCodec.Read(masterCommon, masterContext);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_master_source_background_mismatch",
                        $"Presentation master {masterIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(masterCommon, masterContext))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation master {masterIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(masterCommon, target.Background, masterContext);
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(masterShapeTree, graph.Id, target.Placeholders, masterContext, PartPath(graph.Part)))
                {
                    masterRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, masterContext, changedParts, addedRelationshipIds, addedPartPaths, removedSourceRelationshipKeys, removedSourcePartPaths);
            }

            for (var layoutIndex = 0; layoutIndex < layoutGraph.Length; layoutIndex++)
            {
                var (master, graph) = layoutGraph[layoutIndex];
                var layoutRoot = graph.Part.SlideLayout ??
                    throw new CodecException("missing_slide_layout_root", $"Presentation layout {layoutIndex + 1} has no slide layout root.", PartPath(graph.Part));
                var layoutCommon = layoutRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation layout {layoutIndex + 1} has no common slide data.", PartPath(graph.Part));
                var layoutShapeTree = layoutCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation layout {layoutIndex + 1} has no shape tree.", PartPath(graph.Part));
                var target = envelope.Presentation.Layouts[layoutIndex];
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_layout_binding",
                    $"Presentation layout {layoutIndex + 1} is missing its source binding.",
                    PartPath(graph.Part));
                var sourceName = layoutCommon.Name?.Value ?? $"Layout {graph.Index + 1}";
                if (target.Id != graph.Id || target.MasterId != master.Id || target.Name != sourceName || target.Type != LayoutTypeName(layoutRoot) ||
                    binding.LayoutIndex != graph.Index ||
                    !binding.PartPath.Equals(PartPath(graph.Part), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(graph.RelationshipId, StringComparison.Ordinal) ||
                    !binding.LayoutXmlSha256.Equals(HashElement(layoutRoot), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_binding_mismatch",
                        $"Presentation layout {layoutIndex + 1} does not match its hash-bound read-only source layout.",
                        PartPath(graph.Part));
                var layoutContext = new PptxPartContext(graph.Part, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
                var originalBackground = PptxBackgroundCodec.Read(layoutCommon, layoutContext);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_layout_source_background_mismatch",
                        $"Presentation layout {layoutIndex + 1} background does not match its source binding.",
                        PartPath(graph.Part));
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(layoutCommon, layoutContext))
                        throw new CodecException("unsupported_presentation_edit", $"Presentation layout {layoutIndex + 1} background is preserved but not safely editable by this codec slice.", PartPath(graph.Part));
                    PptxBackgroundCodec.Apply(layoutCommon, target.Background, layoutContext);
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                if (ApplyPlaceholders(layoutShapeTree, graph.Id, target.Placeholders, layoutContext, PartPath(graph.Part)))
                {
                    layoutRoot.Save();
                    changedParts.Add(PartPath(graph.Part));
                }
                TrackContextChanges(graph.Part, layoutContext, changedParts, addedRelationshipIds, addedPartPaths, removedSourceRelationshipKeys, removedSourcePartPaths);
            }

            ulong semanticItems = 0;
            for (var targetPosition = 0; targetPosition < retainedTargets.Length; targetPosition++)
            {
                var targetSlide = retainedTargets[targetPosition];
                var slideIndex = targetSlide.TargetIndex;
                var relationshipId = targetSlide.Source.RelationshipId;
                var slidePart = targetSlide.Source.Part;
                var slideRoot = slidePart.Slide ??
                    throw new CodecException("missing_slide_root", $"Presentation slide {slideIndex + 1} has no slide root.", PartPath(slidePart));
                var target = targetSlide.Target;
                var binding = target.Source ?? throw new CodecException(
                    "missing_presentation_slide_binding",
                    $"Presentation slide {slideIndex + 1} is missing its source binding.",
                    "ppt/presentation.xml");
                var sourceVisibility = PptxSlideVisibilityCodec.Read(slideRoot);
                if (binding.SlideIndex != targetSlide.Source.Index ||
                    !binding.PartPath.Equals(PartPath(slidePart), StringComparison.OrdinalIgnoreCase) ||
                    !binding.RelationshipId.Equals(relationshipId, StringComparison.Ordinal) ||
                    !binding.SlideXmlSha256.Equals(HashElement(slideRoot), StringComparison.OrdinalIgnoreCase) ||
                    binding.SpeakerNotesAddable != PptxSpeakerNotesCodec.CanAddSourceBound(presentationPart, slidePart) ||
                    binding.LegacyCommentsAddable != PptxLegacyCommentsCodec.CanAddSourceBound(presentationPart, slidePart) ||
                    binding.LegacyCommentsEditable != PptxLegacyCommentsCodec.CanEditSourceBound(presentationPart, slidePart, targetSlide.Source.Index) ||
                    binding.CommentPartPresent != PptxLegacyCommentsCodec.CommentPartPresent(slidePart) ||
                    !binding.CommentFamily.Equals(PptxLegacyCommentsCodec.CommentFamily(presentationPart), StringComparison.Ordinal) ||
                    binding.TransitionEditable != PptxTransitionCodec.Supports(slideRoot) ||
                    binding.TransitionPresent != PptxTransitionCodec.HasTransition(slideRoot) ||
                    binding.TransitionAddable != PptxTransitionCodec.CanAdd(slideRoot) ||
                    binding.VisibilityEditable != sourceVisibility.Editable ||
                    !binding.VisibilitySemanticSha256.Equals(sourceVisibility.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} does not match its hash-bound source slide.",
                        PartPath(slidePart));
                var sourceLayoutPart = slidePart.SlideLayoutPart;
                var sourceLayoutId = sourceLayoutPart is null ? string.Empty :
                    layoutIdByPartPath.GetValueOrDefault(PartPath(sourceLayoutPart)) ??
                    throw new CodecException("unresolved_slide_layout_binding", $"Presentation slide {slideIndex + 1} references a layout outside the master graph.", PartPath(slidePart));
                var sourceLayoutRelationshipId = sourceLayoutPart is null ? string.Empty : slidePart.GetIdOfPart(sourceLayoutPart);
                if (target.LayoutId != sourceLayoutId || binding.LayoutRelationshipId != sourceLayoutRelationshipId)
                    throw new CodecException(
                        "presentation_slide_layout_binding_changed",
                        $"Source-preserving PPTX export cannot change slide {slideIndex + 1}'s layout binding.",
                        PartPath(slidePart));

                var slideCommon = slideRoot.CommonSlideData ??
                    throw new CodecException("missing_common_slide_data", $"Presentation slide {slideIndex + 1} has no common slide data.", PartPath(slidePart));
                var shapeTree = slideCommon.ShapeTree ??
                    throw new CodecException("missing_shape_tree", $"Presentation slide {slideIndex + 1} has no shape tree.", PartPath(slidePart));
                var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog, slideNumber: checked(slideIndex + 1));
                var originalBackground = PptxBackgroundCodec.Read(slideCommon, slideContext);
                var originalBackgroundHash = BackgroundSemanticHash(originalBackground);
                if (!binding.BackgroundSemanticSha256.Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_background_mismatch",
                        $"Presentation slide {slideIndex + 1} background does not match its source binding.",
                        PartPath(slidePart));
                var changed = false;
                if (sourceVisibility.Editable != target.HasHidden)
                    throw new CodecException(
                        "presentation_slide_visibility_binding_mismatch",
                        $"Presentation slide {slideIndex + 1} visibility no longer matches its source capability.",
                        PartPath(slidePart));
                if (target.HasHidden && target.Hidden != sourceVisibility.Hidden && PptxSlideVisibilityCodec.ApplySourceBound(slideRoot, target))
                    changed = true;
                var sourceName = slideCommon.Name?.Value ?? string.Empty;
                if (!string.Equals(target.Name, sourceName, StringComparison.Ordinal))
                {
                    // This is deliberately the only source-bound slide metadata
                    // mutation: p:cSld/@name belongs to the existing SlidePart
                    // and does not alter its relationship graph or shape tree.
                    slideCommon.Name = target.Name;
                    changed = true;
                }
                if (!BackgroundSemanticHash(target.Background).Equals(originalBackgroundHash, StringComparison.OrdinalIgnoreCase))
                {
                    if (!binding.BackgroundEditable || !PptxBackgroundCodec.Supports(slideCommon, slideContext))
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} background is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxBackgroundCodec.Apply(slideCommon, target.Background, slideContext);
                    changed = true;
                }
                var originalTransition = PptxTransitionCodec.Read(slideRoot);
                var originalTransitionHash = PptxTransitionCodec.SemanticHash(originalTransition);
                if (!binding.TransitionSemanticSha256.Equals(originalTransitionHash, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_transition_mismatch",
                        $"Presentation slide {slideIndex + 1} transition does not match its source binding.",
                        PartPath(slidePart));
                if (!PptxTransitionCodec.SemanticHash(target.Transition).Equals(originalTransitionHash, StringComparison.OrdinalIgnoreCase))
                {
                    var canAddTransition = binding.TransitionAddable && PptxTransitionCodec.CanAdd(slideRoot);
                    if ((!binding.TransitionEditable || !PptxTransitionCodec.Supports(slideRoot)) && !canAddTransition)
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} transition is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxTransitionCodec.Apply(slideRoot, target.Transition);
                    changed = true;
                }
                var sourceElements = ShapeElements(shapeTree);
                var zOrderPlan = AnalyzeElementZOrder(sourceElements);
                // Timing edits can change whether a native shape is safe to
                // delete, but the imported element binding is about the
                // source revision. Capture that contract before replacing the
                // timing tree so an animation edit does not masquerade as a
                // shape-capability mutation.
                var deletionAnalysis = PptxElementDeletionCodec.AnalyzeSlide(slidePart);
                var sourceDeletionPlans = sourceElements
                    .Select(element => PptxElementDeletionCodec.Analyze(slidePart, element, sourceElements, deletionAnalysis))
                    .ToArray();
                var (retainedElements, authoredElements) = SplitSourceBoundElements(target, sourceElements.Length, slideIndex, slidePart);
                var elementIdsByNativeId = NativeElementIds(sourceElements, target.Id);
                var nativeIdsByElementId = elementIdsByNativeId.ToDictionary(item => item.Value, item => item.Key, StringComparer.Ordinal);
                P.Slide? previousSourceRoot = null;
                IReadOnlyDictionary<uint, string>? previousSourceElementIds = null;
                string? previousSourceSlideId = null;
                if (targetSlide.Source.Index > 0 && PptxTimingCodec.HasMorph(slideRoot))
                {
                    var previousSourceTarget = retainedTargets.FirstOrDefault(candidate =>
                        candidate.Source.Index == targetSlide.Source.Index - 1);
                    if (previousSourceTarget is null)
                        throw new CodecException(
                            "invalid_presentation_morph",
                            $"Presentation slide {slideIndex + 1} cannot retain Morph after its source predecessor was removed.",
                            PartPath(slidePart));
                    previousSourceRoot = previousSourceTarget.Source.Part.Slide;
                    var previousSourceShapeTree = previousSourceRoot?.CommonSlideData?.ShapeTree;
                    previousSourceSlideId = previousSourceTarget.Target.Id;
                    if (previousSourceShapeTree is not null)
                        previousSourceElementIds = NativeElementIds(ShapeElements(previousSourceShapeTree), previousSourceSlideId);
                }
                var originalTiming = PptxTimingCodec.Read(
                    slideRoot,
                    elementIdsByNativeId,
                    previousSourceRoot,
                    previousSourceElementIds,
                    previousSourceSlideId);
                if (!binding.TimingSemanticSha256.Equals(originalTiming.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_slide_source_timing_mismatch",
                        $"Presentation slide {slideIndex + 1} timing does not match its source binding.",
                        PartPath(slidePart));
                PptxTimingCodec.ValidateMorphContext(target, slideIndex > 0 ? targetSlides[slideIndex - 1].Target : null);
                var requestedTimingHash = PptxTimingCodec.SemanticHash(target.Animations, target.Morph);
                var requestedOpaqueNoop = originalTiming.Present && !originalTiming.Editable && target.Animations.Count == 0 && target.Morph is null;
                if (!requestedOpaqueNoop && !requestedTimingHash.Equals(originalTiming.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                {
                    var canReplaceTiming = originalTiming.Editable || (originalTiming.Addable && (target.Animations.Count > 0 || target.Morph is not null));
                    if (!canReplaceTiming)
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            $"Presentation slide {slideIndex + 1} timing is preserved but not safely editable by this codec slice.",
                            PartPath(slidePart));
                    PptxTimingCodec.Apply(slideRoot, target, nativeIdsByElementId, allowOpaqueReplacement: originalTiming.Editable || originalTiming.Addable);
                    changed = true;
                }
                foreach (var (elementId, nativeId) in AuthoredOverlayNativeIds(sourceElements, authoredElements, slideIndex, slidePart))
                    if (!nativeIdsByElementId.TryAdd(elementId, nativeId))
                        throw new CodecException(
                            "invalid_presentation_element",
                            $"Presentation slide {slideIndex + 1} authored overlay element {elementId} reuses an existing source element identity.",
                            PartPath(slidePart));

                var requestedBySourceIndex = new Dictionary<int, PresentationElement>();
                foreach (var requested in retainedElements)
                {
                    var requestedBinding = requested.Source ?? throw new CodecException(
                        "missing_presentation_element_binding",
                        $"Presentation slide {slideIndex + 1} retained element {requested.Id} is missing its source binding.",
                        PartPath(slidePart));
                    if (requestedBinding.ShapeTreeIndex >= (uint)sourceElements.Length)
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} retained element {requested.Id} identifies a source shape-tree index outside the slide.",
                            PartPath(slidePart));
                    var sourceIndex = (int)requestedBinding.ShapeTreeIndex;
                    if (!requestedBySourceIndex.TryAdd(sourceIndex, requested))
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} retained elements must identify unique source shape-tree nodes.",
                            PartPath(slidePart));
                }
                var deletionsBySourceIndex = new Dictionary<int, PresentationElementDeletion>();
                foreach (var deletion in target.ElementDeletions)
                {
                    var deletionBinding = deletion.Source ?? throw new CodecException(
                        "missing_presentation_element_deletion_binding",
                        $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} is missing its source binding.",
                        PartPath(slidePart));
                    if (deletionBinding.ShapeTreeIndex >= (uint)sourceElements.Length)
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} identifies a source shape-tree index outside the slide.",
                            PartPath(slidePart));
                    var sourceIndex = (int)deletionBinding.ShapeTreeIndex;
                    if (requestedBySourceIndex.ContainsKey(sourceIndex) || !deletionsBySourceIndex.TryAdd(sourceIndex, deletion))
                        throw new CodecException(
                            "presentation_element_topology_changed",
                            $"Presentation slide {slideIndex + 1} element deletion {deletion.Id} does not identify one omitted source element.",
                            PartPath(slidePart));
                }
                var requestedSourceOrder = retainedElements
                    .Select(element => checked((int)element.Source!.ShapeTreeIndex))
                    .ToArray();
                var retainedSourceOrder = Enumerable.Range(0, sourceElements.Length)
                    .Where(index => !deletionsBySourceIndex.ContainsKey(index));
                var sourceOrderChanged = !requestedSourceOrder.SequenceEqual(retainedSourceOrder);
                if (sourceOrderChanged && (authoredElements.Length > 0 || deletionsBySourceIndex.Count > 0))
                    throw new CodecException(
                        "unsupported_presentation_element_reorder",
                        $"Presentation slide {slideIndex + 1} cannot combine direct-element reordering with authored overlays or deletions in one export; commit and reopen between bounded edits.",
                        PartPath(slidePart));
                if (sourceOrderChanged && !zOrderPlan.Supported)
                    throw new CodecException(
                        "unsupported_presentation_element_reorder",
                        $"Presentation slide {slideIndex + 1} cannot safely reorder its direct elements: {zOrderPlan.BlockedReason}.",
                        PartPath(slidePart));
                var pendingElementDeletions = new List<(OpenXmlElement Source, PptxElementDeletionPlan Plan)>();
                for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
                {
                    semanticItems++;
                    if (semanticItems > limits.MaxCells)
                        throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    var sourceElement = sourceElements[elementIndex];
                    // Element identities are owned by the source-bound slide, not
                    // by its current presentation position. A slide reorder must
                    // therefore retain the imported owner ID while proving each
                    // element binding against the original SlidePart.
                    var original = ReadElement(sourceElement, target.Id, elementIndex, slideContext, nativeObjects, elementIdsByNativeId);
                    var deletionPlan = sourceDeletionPlans[elementIndex];
                    SetElementDeletionCapability(original, deletionPlan);
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Table)
                    {
                        semanticItems += checked((ulong)original.Table.Rows.Sum(row => row.Cells.Count));
                        if (semanticItems > limits.MaxCells)
                            throw new CodecException("presentation_item_budget_exceeded", $"PPTX presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).", PartPath(slidePart));
                    }
                    if (!requestedBySourceIndex.TryGetValue(elementIndex, out var requested))
                    {
                        if (!deletionsBySourceIndex.TryGetValue(elementIndex, out var deletion))
                            throw new CodecException(
                                "presentation_element_topology_changed",
                                $"Presentation slide {slideIndex + 1} source element {elementIndex + 1} is neither retained nor explicitly deleted.",
                                PartPath(slidePart));
                        SetElementZOrderCapability(original, zOrderPlan);
                        AssertElementBinding(deletion.Id, deletion.Source, sourceElement, original, deletionPlan, zOrderPlan, slideIndex, elementIndex, slidePart);
                        if (!deletion.Id.Equals(original.Id, StringComparison.Ordinal))
                            throw new CodecException(
                                "presentation_element_deletion_binding_mismatch",
                                $"Presentation slide {slideIndex + 1} deletion {elementIndex + 1} changed its source element identity.",
                                PartPath(slidePart));
                        if (!deletionPlan.Supported)
                            throw new CodecException(
                                "unsupported_presentation_element_delete",
                                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} cannot be safely deleted: {deletionPlan.BlockedReason}.",
                                PartPath(slidePart));
                        pendingElementDeletions.Add((sourceElement, deletionPlan));
                        continue;
                    }
                    var elementBinding = requested.Source!;
                    SetElementZOrderCapability(original, zOrderPlan);
                    AssertElementBinding(requested.Id, elementBinding, sourceElement, original, deletionPlan, zOrderPlan, slideIndex, elementIndex, slidePart);
                    PptxOleWorkbookReplacement? oleWorkbookReplacement = null;
                    PptxOleOfficePackageReplacement? oleOfficePackageReplacement = null;
                    PptxDiagramTextReplacement? diagramTextReplacement = null;
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                        PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        ValidateNativeObjectRequest(original, requested);
                        oleWorkbookReplacement = PptxOleWorkbookCodec.PrepareReplacement(original.Opaque, requested.Opaque, assetCatalog, limits);
                        oleOfficePackageReplacement = PptxOleOfficePackageCodec.PrepareReplacement(original.Opaque, requested.Opaque, assetCatalog, limits);
                        diagramTextReplacement = PptxDiagramTextCodec.PrepareReplacement(slidePart, sourceElement, original.Opaque, requested.Opaque);
                    }
                    if (!HasParagraphTabStopRemovalIntent(requested) && !HasParagraphRightMarginRemovalIntent(requested) && SemanticHash(requested).Equals(elementBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                    var stateChanged = PptxElementStateCodec.StateChanged(original, requested);
                    var contentChanged = !PptxElementStateCodec.EqualExceptState(original, requested);
                    if (stateChanged)
                    {
                        PptxElementStateCodec.ApplyBound(sourceElement, original, requested);
                        changed = true;
                    }
                    if (!contentChanged) continue;
                    if (IsMediaAccessibilityOnlyChange(sourceElement, original, requested))
                    {
                        ApplyMediaAccessibility(sourceElement, requested);
                        changed = true;
                        continue;
                    }
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Diagram &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Group &&
                        original.Diagram.DrawingCacheVerified &&
                        original.Diagram.Drawing.Equals(requested.Group))
                    {
                        if (!deletionPlan.Supported)
                            throw new CodecException(
                                "unsupported_presentation_smartart_detach",
                                $"Presentation slide {slideIndex + 1} SmartArt cannot detach because its native closure is not exclusively owned: {deletionPlan.BlockedReason}.",
                                PartPath(slidePart));
                        var nextNativeId = nativeIdsByElementId.Values.DefaultIfEmpty(1u).Max() + 1u;
                        foreach (var detachedElement in FlattenPresentationElements([requested]))
                            if (!nativeIdsByElementId.ContainsKey(detachedElement.Id))
                                nativeIdsByElementId.Add(detachedElement.Id, checked(nextNativeId++));
                        sourceElement.InsertBeforeSelf(BuildElement(requested, nativeIdsByElementId, slideContext, slidePart));
                        PptxElementDeletionCodec.Apply(slidePart, sourceElement, deletionPlan);
                        foreach (var removedRelationshipId in deletionPlan.RelationshipIds)
                        {
                            removedSourceRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                            removedElementRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                        }
                        if (deletionPlan.RelationshipIds.Count > 0)
                            changedParts.Add(RelationshipPartPath(slidePart));
                        if (deletionPlan.RemovedPackagePartPaths.Count > 0)
                        {
                            changedParts.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            removedSourcePartPaths.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            removedElementPartPaths.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            changedParts.Add("[Content_Types].xml");
                        }
                        changed = true;
                        continue;
                    }
                    if (original.ContentCase == PresentationElement.ContentOneofCase.Diagram &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Diagram)
                    {
                        if (!deletionPlan.Supported)
                            throw new CodecException(
                                "unsupported_presentation_smartart_edit",
                                $"Presentation slide {slideIndex + 1} SmartArt cannot be rewritten because its native closure is not exclusively owned: {deletionPlan.BlockedReason}.",
                                PartPath(slidePart));
                        sourceElement.InsertBeforeSelf(BuildElement(requested, nativeIdsByElementId, slideContext, slidePart));
                        PptxElementDeletionCodec.Apply(slidePart, sourceElement, deletionPlan);
                        foreach (var removedRelationshipId in deletionPlan.RelationshipIds)
                        {
                            removedSourceRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                            removedElementRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                        }
                        if (deletionPlan.RelationshipIds.Count > 0)
                            changedParts.Add(RelationshipPartPath(slidePart));
                        if (deletionPlan.RemovedPackagePartPaths.Count > 0)
                        {
                            changedParts.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            removedSourcePartPaths.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            removedElementPartPaths.UnionWith(deletionPlan.RemovedPackagePartPaths);
                            changedParts.Add("[Content_Types].xml");
                        }
                        changed = true;
                        continue;
                    }
                    if (!elementBinding.Editable)
                    {
                        if (elementBinding.TextEditable &&
                            sourceElement is P.Shape sourcePlaceholder &&
                            requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                            PptxPlaceholderCodec.SupportsSlideTextEditing(sourcePlaceholder))
                        {
                            PptxPlaceholderCodec.ApplySlideText(sourcePlaceholder, original, requested, slideContext);
                            changed = true;
                            continue;
                        }
                        if (elementBinding.TextEditable &&
                            sourceElement is P.Shape sourceTextShape &&
                            IsBoundTextOnlyShapeEdit(original, requested, sourceTextShape))
                        {
                            PptxTextCodec.Apply(sourceTextShape, requested.Shape, slideContext);
                            changed = true;
                            continue;
                        }
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                    if (sourceElement is P.Shape sourceShape &&
                        requested.ContentCase == PresentationElement.ContentOneofCase.Shape &&
                        IsSimpleShape(sourceShape, slideContext, allowNegativeOffset: Geometry(sourceShape) == "line"))
                    {
                        ApplyShape(sourceShape, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.Picture sourcePicture &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Image &&
                             PptxPictureCodec.TryRead(sourcePicture, slideContext, out _))
                    {
                        PptxPictureCodec.Apply(sourcePicture, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceTable &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Table &&
                             PptxTableCodec.TryRead(sourceTable, slideContext, out _))
                    {
                        PptxTableCodec.Apply(sourceTable, requested, slideContext);
                        changed = true;
                    }
                    else if (sourceElement is P.ConnectionShape sourceConnector &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Connector &&
                             PptxConnectorCodec.TryRead(sourceConnector, elementIdsByNativeId, out _))
                    {
                        PptxConnectorCodec.Apply(sourceConnector, requested, nativeIdsByElementId);
                        changed = true;
                    }
                    else if (sourceElement is P.GraphicFrame sourceChart &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Chart &&
                             PptxChartCodec.TryRead(sourceChart, slideContext, out _, out var chartEditable) && chartEditable)
                    {
                        var replacement = PptxChartCodec.Apply(sourceChart, requested, slideContext, limits);
                        changedParts.Add(replacement.PartPath);
                        changedParts.UnionWith(replacement.ChangedPartPaths);
                        addedRelationshipIds.UnionWith(replacement.AddedRelationshipKeys);
                        addedPartPaths.UnionWith(replacement.AddedPartPaths);
                        removedSourceRelationshipKeys.UnionWith(replacement.RemovedRelationshipKeys);
                        removedSourcePartPaths.UnionWith(replacement.RemovedPartPaths);
                        replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
                        foreach (var entry in replacement.ReplacedPartHashes) replacedOpaquePartHashes.Add(entry.Key, entry.Value);
                        changed |= replacement.SlideChanged;
                    }
                    else if (sourceElement is P.GroupShape sourceGroup &&
                             requested.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             original.ContentCase == PresentationElement.ContentOneofCase.Group &&
                             TryReadGroup(sourceGroup, original.Id, slideContext, elementIdsByNativeId, out _))
                    {
                        if (ApplyGroup(sourceGroup, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"element {elementIndex + 1}", limits))
                            changed = true;
                    }
                    else if (requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                             PptxNativeObjectCatalog.SupportsPlacementEditing(sourceElement))
                    {
                        if (oleWorkbookReplacement is not null)
                        {
                            PptxOleWorkbookCodec.Apply(slidePart, sourceElement, original.Opaque.OleWorkbook, oleWorkbookReplacement);
                            changedParts.Add(oleWorkbookReplacement.PartPath);
                            replacedOpaquePartHashes.Add(oleWorkbookReplacement.PartPath, oleWorkbookReplacement.Sha256);
                        }
                        if (oleOfficePackageReplacement is not null)
                        {
                            PptxOleOfficePackageCodec.Apply(slidePart, sourceElement, original.Opaque.OleOfficePackage, oleOfficePackageReplacement);
                            changedParts.Add(oleOfficePackageReplacement.PartPath);
                            replacedOpaquePartHashes.Add(oleOfficePackageReplacement.PartPath, oleOfficePackageReplacement.Sha256);
                        }
                        if (diagramTextReplacement is not null)
                        {
                            PptxDiagramTextCodec.Apply(slidePart, original.Opaque.DiagramText, diagramTextReplacement);
                            changedParts.Add(diagramTextReplacement.PartPath);
                            replacedOpaquePartHashes.Add(diagramTextReplacement.PartPath, diagramTextReplacement.Sha256);
                        }
                        if (NativePlacementChanged(original, requested))
                        {
                            ApplyNativePlacement(sourceElement, requested);
                            changed = true;
                        }
                    }
                    else
                    {
                        throw UnsupportedPresentationEdit(slideIndex, elementIndex, slidePart);
                    }
                }
                if (sourceOrderChanged)
                {
                    ApplyElementZOrder(shapeTree, sourceElements, requestedSourceOrder);
                    changed = true;
                }
                foreach (var deletion in pendingElementDeletions)
                {
                    PptxElementDeletionCodec.Apply(slidePart, deletion.Source, deletion.Plan);
                    foreach (var removedRelationshipId in deletion.Plan.RelationshipIds)
                    {
                        removedSourceRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                        removedElementRelationshipKeys.Add($"{PartPath(slidePart)}\0{removedRelationshipId}");
                    }
                    if (deletion.Plan.RelationshipIds.Count > 0)
                        changedParts.Add(RelationshipPartPath(slidePart));
                    if (deletion.Plan.RemovedPackagePartPaths.Count > 0)
                    {
                        changedParts.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        removedSourcePartPaths.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        removedElementPartPaths.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                        changedParts.Add("[Content_Types].xml");
                    }
                    changed = true;
                }
                if (authoredElements.Length > 0)
                {
                    if (changed)
                        throw new CodecException(
                            "unsupported_presentation_authored_overlay",
                            $"Presentation slide {slideIndex + 1} cannot combine an authored overlay with another SlidePart mutation in one export; commit and reopen between bounded edits.",
                            PartPath(slidePart));
                    var relationshipCount = slideContext.AddedRelationshipIds.Count;
                    var partCount = slideContext.AddedPartPaths.Count;
                    var authoredXml = authoredElements
                        .Select(authored => BuildElement(authored, nativeIdsByElementId, slideContext, slidePart).OuterXml)
                        .ToArray();
                    var authoredImageCount = authoredElements.Count(element => element.ContentCase == PresentationElement.ContentOneofCase.Image);
                    var addedRelationshipCount = slideContext.AddedRelationshipIds.Count - relationshipCount;
                    var addedPartCount = slideContext.AddedPartPaths.Count - partCount;
                    if (addedRelationshipCount < 0 || addedPartCount < 0 ||
                        addedRelationshipCount > authoredImageCount || addedPartCount > addedRelationshipCount)
                        throw new CodecException(
                            "unsupported_presentation_authored_overlay",
                            $"Presentation slide {slideIndex + 1} authored overlay changed relationships outside its embedded-image allowance.",
                            PartPath(slidePart));
                    authoredOverlayXmlByPartPath.Add(PartPath(slidePart), authoredXml);
                    changedParts.Add(PartPath(slidePart));
                }
                if (changed)
                {
                    slideRoot.Save();
                    changedParts.Add(PartPath(slidePart));
                }
                if (PptxSpeakerNotesCodec.ApplySourceBound(presentationPart, slidePart, target.SpeakerNotes, slideIndex) is { } notesChange)
                {
                    changedParts.UnionWith(notesChange.ChangedPartPaths);
                    addedPartPaths.UnionWith(notesChange.AddedPartPaths);
                    addedRelationshipIds.UnionWith(notesChange.AddedRelationshipKeys);
                    foreach (var (partPath, sha256) in notesChange.ReplacedPartHashes)
                        replacedOpaquePartHashes.Add(partPath, sha256);
                }
                if (PptxModernCommentsCodec.ApplySourceBound(
                        presentationPart,
                        targetSlide.Source.SlideId,
                        slidePart,
                        sourceElements,
                        elementIdsByNativeId,
                        target,
                        slideIndex) is { } modernCommentsChange)
                {
                    changedParts.Add(modernCommentsChange.PartPath);
                    replacedOpaquePartHashes.Add(modernCommentsChange.PartPath, modernCommentsChange.Sha256);
                }
                TrackContextChanges(slidePart, slideContext, changedParts, addedRelationshipIds, addedPartPaths, removedSourceRelationshipKeys, removedSourcePartPaths);
            }
            if (PptxLegacyCommentsCodec.ApplySourceBoundEdits(
                    presentationPart,
                    slideParts,
                    retainedTargets.Select(target => target.Target).ToArray()) is { } legacyCommentsEdit)
            {
                changedParts.UnionWith(legacyCommentsEdit.ChangedPartPaths);
                addedPartPaths.UnionWith(legacyCommentsEdit.AddedPartPaths);
                addedRelationshipIds.UnionWith(legacyCommentsEdit.AddedRelationshipKeys);
                foreach (var (partPath, sha256) in legacyCommentsEdit.ReplacedPartHashes)
                    replacedOpaquePartHashes.Add(partPath, sha256);
            }
            if (PptxLegacyCommentsCodec.ApplySourceBoundAdditions(
                    presentationPart,
                    slideParts,
                    retainedTargets.Select(target => target.Target).ToArray()) is { } legacyCommentsChange)
            {
                changedParts.UnionWith(legacyCommentsChange.ChangedPartPaths);
                addedPartPaths.UnionWith(legacyCommentsChange.AddedPartPaths);
                addedRelationshipIds.UnionWith(legacyCommentsChange.AddedRelationshipKeys);
                foreach (var (partPath, sha256) in legacyCommentsChange.ReplacedPartHashes)
                    replacedOpaquePartHashes.Add(partPath, sha256);
            }
        }

        // Opening an OPC package can rewrite ZIP container metadata even with
        // AutoSave disabled. When no modeled part, relationship, or opaque
        // graph changed, the validated source bytes are the only lossless
        // export: returning the stream would create a spurious package-level
        // diff while every retained part remains identical.
        var noModeledChanges = changedParts.Count == 0 &&
            addedRelationshipIds.Count == 0 &&
            addedPartPaths.Count == 0 &&
            replacedOpaquePartHashes.Count == 0 &&
            removedSourcePartPaths.Count == 0 &&
            removedSourceRelationshipKeys.Count == 0;
        var bytes = noModeledChanges
            ? sourceBytes
            : NormalizeChangedPartTimestamps(
                stream,
                sourceBytes,
                addedPartPaths.Concat(clonedPackageEntryPaths).ToHashSet(StringComparer.OrdinalIgnoreCase),
                changedParts);
        if (authoredOverlayXmlByPartPath.Count > 0)
        {
            var replacements = authoredOverlayXmlByPartPath.ToDictionary(
                entry => entry.Key,
                entry => PptxEditPlanCodec.AppendShapeTreeChildren(
                    PptxEditPlanCodec.ReadPart(bytes, entry.Key),
                    entry.Value,
                    entry.Key),
                StringComparer.OrdinalIgnoreCase);
            bytes = PptxEditPlanCodec.ReplaceParts(bytes, replacements);
        }
        // The normalized output owns its exact byte array now. Drop the
        // expandable ZIP work buffer before opening source/output packages for
        // post-write validation so both large buffers are not live together.
        stream.SetLength(0);
        stream.Capacity = 0;
        IReadOnlyList<string> actualChangedParts;
        int retainedValidationErrorCount;
        using (PpjBuildProfiler.Measure("post-write.validation"))
        {
            ValidateOutputBudget(bytes, limits);
            AssertPlannedPartsRemoved(sourceBytes, bytes, removedSourcePartPaths);
            retainedValidationErrorCount = ValidateOffice2021AgainstSource(sourceBytes, bytes, clonedPartSourcePaths);
            actualChangedParts = AssertPackagePartsUnchangedExcept(sourceBytes, bytes, changedParts);
            var actualChangedPartSet = actualChangedParts.ToHashSet(StringComparer.OrdinalIgnoreCase);
            ValidatePreservedSlideElements(sourceBytes, bytes, envelope.Presentation, limits, actualChangedPartSet);
            ValidatePreservedMasterAndLayoutContent(sourceBytes, bytes, envelope.Presentation, limits, actualChangedPartSet);
            var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
            AssertOpaqueGraphMatchesWithModeledAdditions(
                envelope.OpaqueOpc,
                outputOpaque,
                addedRelationshipIds,
                addedPartPaths,
                replacedOpaquePartHashes,
                removedSourcePartPaths,
                removedSourceRelationshipKeys);
        }
        var diagnostics = new List<Diagnostic>();
        var removedElementOpaqueCount = envelope.OpaqueOpc.Parts.Count(part => removedElementPartPaths.Contains(part.Path)) +
                                        envelope.OpaqueOpc.PackageRelationships.Count(relationship =>
                                            removedElementPartPaths.Contains(relationship.SourcePath) ||
                                            removedElementRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"));
        var removedOpaqueCount = envelope.OpaqueOpc.Parts.Count(part => removedSourcePartPaths.Contains(part.Path)) +
                                 envelope.OpaqueOpc.PackageRelationships.Count(relationship =>
                                     removedSourcePartPaths.Contains(relationship.SourcePath) ||
                                     removedSourceRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"));
        var removedSlideOpaqueCount = removedOpaqueCount - removedElementOpaqueCount;
        var retainedOpaqueCount = opaqueCount - removedOpaqueCount;
        if (retainedOpaqueCount > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_content_preserved",
                $"Preserved {retainedOpaqueCount} opaque OPC parts or relationships while updating modeled presentation content."));
        if (removedSlideOpaqueCount > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_content_deleted_with_slide",
                $"Removed {removedSlideOpaqueCount} opaque OPC parts or relationships because they belonged exclusively to an explicitly deleted source slide graph."));
        if (removedElementOpaqueCount > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
                "opaque_content_deleted_with_element",
                $"Removed {removedElementOpaqueCount} opaque OPC parts or relationships because they belonged exclusively to explicitly deleted source slide elements."));
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecDiagnostics.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; export introduced none."));
        return new PptxExportResult(bytes, diagnostics, actualChangedParts);
    }

    private static byte[] NormalizeChangedPartTimestamps(
        MemoryStream stream,
        byte[] sourceBytes,
        IReadOnlyCollection<string> addedPartPaths,
        IReadOnlyCollection<string> changedPartPaths)
    {
        if (addedPartPaths.Count == 0 && changedPartPaths.Count == 0) return stream.ToArray();
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var timestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            var added = addedPartPaths.ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var path in added.Order(StringComparer.OrdinalIgnoreCase))
            {
                var entry = archive.GetEntry(path) ?? throw new CodecException(
                    "presentation_added_part_missing",
                    $"PPTX declared added part {path} is missing from the output package.",
                    path);
                entry.LastWriteTime = timestamp;
            }
            foreach (var path in changedPartPaths.Where(path => !added.Contains(path)).Order(StringComparer.OrdinalIgnoreCase))
            {
                var outputEntry = archive.GetEntry(path);
                var sourceEntry = sourceArchive.GetEntry(path);
                if (outputEntry is not null && sourceEntry is not null)
                    outputEntry.LastWriteTime = sourceEntry.LastWriteTime;
            }
        }
        return stream.ToArray();
    }

    // This is deliberately a canvas-only mutation. PresentationML leaves all
    // slide, layout, and master coordinates untouched when p:sldSz changes;
    // callers that want a reflow must explicitly use their layout primitives.
    // Clearing the preset type avoids claiming that an arbitrary pair of EMU
    // dimensions still matches the old Office preset.
    private static bool ApplySourceBoundSlideSize(PresentationPart presentationPart, PresentationArtifact requested)
    {
        var presentation = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var source = presentation.SlideSize;
        var sourceWidth = source?.Cx?.Value ?? DefaultSlideWidthEmu;
        var sourceHeight = source?.Cy?.Value ?? DefaultSlideHeightEmu;
        if (requested.SlideWidthEmu == sourceWidth && requested.SlideHeightEmu == sourceHeight)
            return false;
        if (requested.SlideWidthEmu <= 0 || requested.SlideHeightEmu <= 0 ||
            requested.SlideWidthEmu > int.MaxValue || requested.SlideHeightEmu > int.MaxValue)
            throw new CodecException(
                "invalid_slide_size",
                "Source-bound PPTX canvas dimensions must be positive signed 32-bit EMU values.",
                "ppt/presentation.xml");

        if (source is null)
        {
            source = new P.SlideSize();
            var slideIdList = presentation.SlideIdList ??
                throw new CodecException("missing_slide_id_list", "PPTX presentation has no slide ID list.", "ppt/presentation.xml");
            presentation.InsertAfter(source, slideIdList);
        }
        source.Cx = checked((int)requested.SlideWidthEmu);
        source.Cy = checked((int)requested.SlideHeightEmu);
        source.Type = null;
        presentation.Save();
        return true;
    }

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        int slideIndex,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null,
        bool allowNegativeOffset = false)
        => ReadElement(source, $"presentation/slide/{slideIndex + 1}", elementIndex, slideContext, nativeObjects, elementIdsByNativeId, allowNegativeOffset);

    private static PresentationElement ReadElement(
        OpenXmlElement source,
        string ownerId,
        int elementIndex,
        PptxPartContext slideContext,
        PptxNativeObjectCatalog? nativeObjects = null,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId = null,
        bool allowNegativeOffset = false)
    {
        var element = new PresentationElement
        {
            Id = $"{ownerId}/element/{elementIndex + 1}",
            Name = ElementName(source, elementIndex),
        };
        var nativeMediaPicture = PptxNativeObjectCatalog.IsMediaPicture(source);
        var editable = false;
        var modeled = false;
        if (source is P.Shape sourceShape)
        {
            // A line can intentionally bleed a few points beyond the canvas
            // (common in imported separator rules). Keep the exception
            // geometry-specific so an unrelated negative frame never widens
            // the ordinary shape editing surface.
            var shapeAllowNegativeOffset = allowNegativeOffset || Geometry(sourceShape) == "line";
            editable = IsSimpleShape(sourceShape, slideContext, shapeAllowNegativeOffset);
            element.Shape = ReadShape(sourceShape, slideContext);
            var nonVisual = sourceShape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            if (PptxHyperlinkCodec.TryReadElementAction(nonVisual, slideContext, out var action))
                element.Shape.Action = action;
            if (PptxHyperlinkCodec.TryReadElementHoverAction(nonVisual, slideContext, out var hoverAction))
                element.Shape.HoverAction = hoverAction;
            modeled = true;
        }
        else if (source is P.Picture sourcePicture && !nativeMediaPicture && PptxPictureCodec.TryRead(sourcePicture, slideContext, out var image))
        {
            element.Image = image;
            editable = true;
            modeled = true;
        }
        else if (source is P.GraphicFrame sourceFrame)
        {
            PresentationDiagram diagram = null!;
            PresentationTable table = null!;
            PresentationChart chart = null!;
            var chartEditable = false;
            var diagramModeled = slideContext.Owner is SlidePart slidePart && slideContext.Assets is { } diagramAssets &&
                PptxSmartArtCodec.TryRead(sourceFrame, slidePart, diagramAssets, out diagram);
            var tableModeled = !diagramModeled && PptxTableCodec.TryRead(sourceFrame, slideContext, out table);
            var chartModeled = !diagramModeled && PptxChartCodec.TryRead(sourceFrame, slideContext, out chart, out chartEditable);
            editable = chartModeled ? chartEditable : tableModeled;
            if (diagramModeled)
                element.Diagram = diagram;
            else if (tableModeled)
                element.Table = table;
            else if (chartModeled)
                element.Chart = chart;
            modeled = diagramModeled || tableModeled || chartModeled;
        }
        else if (source is P.ConnectionShape sourceConnector && PptxConnectorCodec.TryRead(sourceConnector, elementIdsByNativeId, out var connector))
        {
            element.Connector = connector;
            editable = true;
            modeled = true;
        }
        else if (source is P.GroupShape sourceGroup && TryReadGroup(sourceGroup, element.Id, slideContext, elementIdsByNativeId, out var group))
        {
            element.Group = group;
            editable = true;
            modeled = true;
        }
        if (!modeled)
        {
            var frame = ReadFrame(source);
            element.Opaque = new PresentationOpaqueElement
            {
                ElementName = source.LocalName,
                Text = DescendantText(source),
                RawXml = source.OuterXml,
                LeftEmu = frame.Left,
                TopEmu = frame.Top,
                WidthEmu = frame.Width,
                HeightEmu = frame.Height,
            };
            nativeObjects?.Populate(element.Opaque, source, slideContext.Owner);
            if (nativeMediaPicture && source is P.Picture mediaPicture &&
                PptxNonVisualAccessibilityCodec.TryReadResidual(
                    mediaPicture.NonVisualPictureProperties?.NonVisualDrawingProperties,
                    out var mediaAccessibility))
                element.Opaque.Accessibility = mediaAccessibility;
            var nativeKind = string.IsNullOrEmpty(element.Opaque.NativeKind)
                ? PptxNativeObjectCatalog.Classify(source)
                : element.Opaque.NativeKind;
            editable = PptxNativeObjectCatalog.SupportsPlacementEditing(source, nativeKind);
        }
        element.Source = new PresentationElementSourceBinding
        {
            ShapeTreeIndex = checked((uint)elementIndex),
            ElementSha256 = HashElement(source),
            Editable = editable,
            DirectFramePresenceEditable = source is P.Shape slidePlaceholder &&
                PptxPlaceholderCodec.SupportsSlideFrameEditing(slidePlaceholder),
            TextEditable = source is P.Shape textShape && textShape.TextBody is not null && PptxTextCodec.SupportsEditing(textShape.TextBody),
            AccessibilityEditable =
                (element.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                 nativeMediaPicture && source is P.Picture mediaAccessibilityPicture &&
                 PptxNonVisualAccessibilityCodec.SupportsResidual(
                     mediaAccessibilityPicture.NonVisualPictureProperties?.NonVisualDrawingProperties)) ||
                editable && (
                (source is P.Picture accessibilityPicture && element.ContentCase == PresentationElement.ContentOneofCase.Image &&
                 PptxNonVisualAccessibilityCodec.SupportsResidual(accessibilityPicture.NonVisualPictureProperties?.NonVisualDrawingProperties)) ||
                PptxNonVisualAccessibilityCodec.Supports(source switch
                {
                    P.Shape accessibilityShape => accessibilityShape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                    P.ConnectionShape accessibilityConnector => accessibilityConnector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
                    P.GroupShape accessibilityGroup => accessibilityGroup.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
                    P.GraphicFrame accessibilityFrame when element.ContentCase is PresentationElement.ContentOneofCase.Table or PresentationElement.ContentOneofCase.Chart or PresentationElement.ContentOneofCase.Diagram =>
                        accessibilityFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
                    _ => null,
                })),
        };
        PptxElementStateCodec.Populate(source, element);
        element.Source.SemanticSha256 = SemanticHash(element);
        return element;
    }

    private static void SetElementDeletionCapability(
        PresentationElement element,
        PptxElementDeletionPlan plan)
    {
        element.Source.DeletionCapability = new PresentationElementDeletionCapability
        {
            Supported = plan.Supported,
            BlockedReason = plan.BlockedReason,
            NativeId = plan.NativeId,
        };
    }

    private static PresentationElementZOrderCapability AnalyzeElementZOrder(IReadOnlyList<OpenXmlElement> elements)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            if (elements[index] is P.Shape or P.Picture or P.GraphicFrame or P.ConnectionShape or P.GroupShape or P.ContentPart) continue;
            return new PresentationElementZOrderCapability
            {
                Supported = false,
                BlockedReason = $"direct shape-tree child {index + 1} ({elements[index].LocalName}) is not a movable drawing node",
            };
        }
        return new PresentationElementZOrderCapability { Supported = true };
    }

    private static void SetElementZOrderCapability(
        PresentationElement element,
        PresentationElementZOrderCapability plan)
    {
        element.Source.ZOrderCapability = plan.Clone();
    }

    private static void ApplyElementZOrder(
        P.ShapeTree shapeTree,
        IReadOnlyList<OpenXmlElement> sourceElements,
        IReadOnlyList<int> requestedSourceOrder)
    {
        if (sourceElements.Count != requestedSourceOrder.Count || requestedSourceOrder.Distinct().Count() != sourceElements.Count ||
            requestedSourceOrder.Any(index => index < 0 || index >= sourceElements.Count))
            throw new CodecException("presentation_element_topology_changed", "Presentation direct-element reorder must be a complete permutation of the source shape tree.");
        OpenXmlElement anchor = shapeTree.GetFirstChild<P.GroupShapeProperties>() ??
            throw new CodecException("missing_shape_tree", "Presentation shape tree has no group-shape properties anchor.");
        foreach (var sourceElement in sourceElements) sourceElement.Remove();
        foreach (var sourceIndex in requestedSourceOrder)
        {
            var sourceElement = sourceElements[sourceIndex];
            shapeTree.InsertAfter(sourceElement, anchor);
            anchor = sourceElement;
        }
    }

    private static void ApplyGroupElementZOrder(
        P.GroupShape group,
        IReadOnlyList<OpenXmlElement> sourceElements,
        IReadOnlyList<int> requestedSourceOrder)
    {
        if (sourceElements.Count != requestedSourceOrder.Count || requestedSourceOrder.Distinct().Count() != sourceElements.Count ||
            requestedSourceOrder.Any(index => index < 0 || index >= sourceElements.Count))
            throw new CodecException("presentation_group_topology_changed", "Presentation group-child reorder must be a complete permutation of the local source shape tree.");
        if (group.GetFirstChild<P.GroupShapeProperties>() is null)
            throw new CodecException("missing_group_shape_properties", "Presentation group has no group-shape properties anchor.");
        foreach (var sourceElement in sourceElements) sourceElement.Remove();
        foreach (var sourceIndex in requestedSourceOrder)
        {
            var sourceElement = sourceElements[sourceIndex];
            // AppendChild is used for local group children instead of the
            // generic InsertAfter helper: the Open XML SDK can normalize a
            // mixed p:grpSp choice back to its schema order when the anchor
            // is a group-properties sibling.  The group shell remains first;
            // appending only the already-detached drawing nodes preserves
            // their requested local z-order deterministically.
            group.AppendChild(sourceElement);
        }
    }

    private static void AssertElementBinding(
        string requestedId,
        PresentationElementSourceBinding binding,
        OpenXmlElement sourceElement,
        PresentationElement original,
        PptxElementDeletionPlan deletionPlan,
        PresentationElementZOrderCapability zOrderPlan,
        int slideIndex,
        int elementIndex,
        SlidePart slidePart)
    {
        if (!requestedId.Equals(original.Id, StringComparison.Ordinal) ||
            binding.ShapeTreeIndex != elementIndex ||
            !binding.ElementSha256.Equals(HashElement(sourceElement), StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_element_binding_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} does not match its source element.",
                PartPath(slidePart));
        if (binding.Editable != original.Source.Editable ||
            binding.DirectFramePresenceEditable != original.Source.DirectFramePresenceEditable ||
            binding.TextEditable != original.Source.TextEditable ||
            binding.AccessibilityEditable != original.Source.AccessibilityEditable ||
            binding.VisibilityEditable != original.Source.VisibilityEditable ||
            binding.LockingEditable != original.Source.LockingEditable ||
            binding.DeletionCapability is null ||
            binding.DeletionCapability.Supported != deletionPlan.Supported ||
            binding.DeletionCapability.NativeId != deletionPlan.NativeId ||
            !binding.DeletionCapability.BlockedReason.Equals(deletionPlan.BlockedReason, StringComparison.Ordinal) ||
            binding.ZOrderCapability is null ||
            binding.ZOrderCapability.Supported != zOrderPlan.Supported ||
            !binding.ZOrderCapability.BlockedReason.Equals(zOrderPlan.BlockedReason, StringComparison.Ordinal))
            throw new CodecException(
                "presentation_element_binding_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} changed its source capability contract.",
                PartPath(slidePart));
        if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "presentation_source_semantics_mismatch",
                $"Presentation slide {slideIndex + 1} element {elementIndex + 1} source semantics do not match its binding.",
                PartPath(slidePart));
    }

    private static bool TryReadGroup(
        P.GroupShape source,
        string groupId,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string>? elementIdsByNativeId,
        out PresentationGroup group)
    {
        group = new PresentationGroup();
        var nonVisual = source.GetFirstChild<P.NonVisualGroupShapeProperties>();
        var properties = source.GetFirstChild<P.GroupShapeProperties>();
        var transform = properties?.GetFirstChild<A.TransformGroup>();
        if (nonVisual is null || properties is null || transform is null ||
            source.Elements<P.NonVisualGroupShapeProperties>().Count() != 1 ||
            source.Elements<P.GroupShapeProperties>().Count() != 1 ||
            nonVisual.ChildElements.Count != 3 ||
            nonVisual.ChildElements[0] is not P.NonVisualDrawingProperties drawing ||
            nonVisual.ChildElements[1] is not P.NonVisualGroupShapeDrawingProperties groupDrawing ||
            nonVisual.ChildElements[2] is not P.ApplicationNonVisualDrawingProperties application ||
            !SupportsGroupDrawingProperties(groupDrawing) || application.ChildElements.Count != 0 ||
            drawing.Id?.Value is null or 0 || drawing.Name?.Value is not { Length: <= 1_024 } ||
            !HasOnlyAttributes(application) ||
            properties.ChildElements.Count != 1 || properties.FirstChild != transform || !HasOnlyAttributes(properties) ||
            !PptxFrameTransformCodec.TryRead(transform, out var frameTransform) || transform.ChildElements.Count != 4 ||
            transform.ChildElements[0] is not A.Offset offset ||
            transform.ChildElements[1] is not A.Extents extents ||
            transform.ChildElements[2] is not A.ChildOffset childOffset ||
            transform.ChildElements[3] is not A.ChildExtents childExtents ||
            !HasOnlyAttributes(offset, "x", "y") || !HasOnlyAttributes(extents, "cx", "cy") ||
            !HasOnlyAttributes(childOffset, "x", "y") || !HasOnlyAttributes(childExtents, "cx", "cy") ||
            extents.Cx?.Value <= 0 || extents.Cy?.Value <= 0 || childExtents.Cx?.Value <= 0 || childExtents.Cy?.Value <= 0)
            return false;

        group.LeftEmu = offset.X?.Value ?? 0;
        group.TopEmu = offset.Y?.Value ?? 0;
        group.WidthEmu = extents.Cx?.Value ?? 0;
        group.HeightEmu = extents.Cy?.Value ?? 0;
        group.ChildLeftEmu = childOffset.X?.Value ?? 0;
        group.ChildTopEmu = childOffset.Y?.Value ?? 0;
        group.ChildWidthEmu = childExtents.Cx?.Value ?? 0;
        group.ChildHeightEmu = childExtents.Cy?.Value ?? 0;
        group.FrameTransform = frameTransform;
        group.Accessibility = PptxNonVisualAccessibilityCodec.Read(drawing);
        var children = GroupElements(source);
        if (children.Length == 0) return false;
        var zOrderPlan = AnalyzeElementZOrder(children);
        for (var index = 0; index < children.Length; index++)
        {
            var child = ReadElement(
                children[index],
                groupId,
                index,
                slideContext,
                elementIdsByNativeId: elementIdsByNativeId,
                allowNegativeOffset: true);
            // Keep a valid group projected when a child is only partially
            // modeled.  The child remains source-bound and read-only unless a
            // separate capability proves a narrower edit; rejecting the whole
            // group here makes an otherwise safe sibling edit impossible when
            // a vendor extension or unsupported style shares the group.
            if (child.ContentCase is PresentationElement.ContentOneofCase.None ||
                child.Source is null)
                return false;
            SetElementZOrderCapability(child, zOrderPlan);
            group.Children.Add(child);
        }
        return true;
    }

    private static bool SupportsGroupDrawingProperties(P.NonVisualGroupShapeDrawingProperties groupDrawing)
    {
        if (!HasOnlyAttributes(groupDrawing) || groupDrawing.ChildElements.Count == 0) return groupDrawing.ChildElements.Count == 0;
        if (PptxElementStateCodec.RecognizesGroupLockProfile(groupDrawing)) return true;
        if (groupDrawing.ChildElements.Count != 1 || groupDrawing.FirstChild is not A.GroupShapeLocks locks ||
            locks.ChildElements.Count != 0)
            return false;
        // Group locks are UI hints and do not alter the DrawingML render tree.
        // Keep the observed standard noChangeAspect form source-bound while
        // allowing the group's semantic children to remain discoverable.
        return HasOnlyAttributes(locks, "noChangeAspect") &&
            locks.NoChangeAspect?.Value is true;
    }

    private static PresentationShape ReadShape(P.Shape shape, PptxPartContext slideContext)
    {
        var frame = ReadFrame(shape);
        var properties = shape.ShapeProperties;
        var textBody = PptxTextCodec.Read(shape.TextBody, slideContext);
        var placeholder = PptxPlaceholderCodec.ReadIdentity(shape);
        var transform = properties?.Transform2D;
        var geometry = Geometry(shape);
        var solidFill = properties?.GetFirstChild<A.SolidFill>();
        var gradientFill = properties?.GetFirstChild<A.GradientFill>();
        var imageFill = properties?.GetFirstChild<A.BlipFill>();
        _ = PptxGradientFillCodec.TryRead(gradientFill, out var gradientSemantic);
        _ = PptxImagePaintCodec.TryRead(imageFill, slideContext, out var imageSemantic);
        var result = new PresentationShape
        {
            Geometry = geometry,
            LeftEmu = frame.Left,
            TopEmu = frame.Top,
            WidthEmu = frame.Width,
            HeightEmu = frame.Height,
            Text = PptxTextCodec.Flatten(textBody),
            TextBody = textBody,
            FillRgb = PptxColor.SolidRgb(solidFill),
            FillScheme = PptxColor.SolidSchemeWithOpacity(solidFill),
            GradientFill = gradientSemantic.Stops.Count > 0 ? gradientSemantic : null,
            ImageFill = imageSemantic.AssetId.Length > 0 ? imageSemantic : null,
            Placeholder = placeholder,
            DirectFrame = placeholder is null ? null : PptxPlaceholderCodec.ReadDirectFrame(shape),
            Transform = placeholder is null && PptxShapeTransformCodec.Supports(transform, allowSingleZeroExtent: geometry == "line")
                ? PptxShapeTransformCodec.Read(transform!)
                : null,
            Shadow = PptxShadowCodec.TryRead(properties, out var shadow) ? shadow : null,
            Glow = PptxGlowCodec.TryRead(properties, out var glow) ? glow : null,
            InnerShadow = PptxInnerShadowCodec.TryRead(properties, out var innerShadow) ? innerShadow : null,
            Reflection = PptxReflectionCodec.TryRead(properties, out var reflection) ? reflection : null,
            SoftEdge = PptxSoftEdgeCodec.TryRead(properties, out var softEdge) ? softEdge : null,
        };
        if (ReadFillOpacity(solidFill) is { } fillOpacity)
            result.FillOpacityThousandthPercent = fillOpacity;
        PptxLineStyleCodec.ReadForProjection(properties?.GetFirstChild<A.Outline>(), result);
        if (!string.Equals(geometry, "line", StringComparison.Ordinal))
        {
            // Non-line arrowheads remain source-bound. Keep the historic
            // ordinary-shape preview projection valid without implying that
            // those endpoint semantics are editable.
            result.StartArrow = result.EndArrow = result.StartArrowWidth = result.StartArrowLength =
                result.EndArrowWidth = result.EndArrowLength = string.Empty;
        }
        if (shape.UseBackgroundFill?.HasValue == true)
            result.UseBackgroundFill = shape.UseBackgroundFill.Value;
        PptxPresetGeometryAdjustmentCodec.Read(properties?.GetFirstChild<A.PresetGeometry>(), geometry, result);
        if (PptxShapeImageFillCodec.TryRead(properties?.GetFirstChild<A.BlipFill>(), slideContext, out var sourceImageFill))
            result.ImageFillAssetId = sourceImageFill.Id;
        PptxCustomGeometryCodec.Read(properties?.GetFirstChild<A.CustomGeometry>(), frame.Width, frame.Height, result);
        if (properties?.Elements<A.Shape3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Shape3DType>() is { } shape3d)
        {
            if (PptxShape3DCodec.TryReadExtrusionHeight(shape3d, out var extrusionHeight))
                result.ShapeExtrusionHeightEmu = extrusionHeight;
            if (PptxShape3DCodec.TryReadDepth(shape3d, out var depth))
                result.ShapeDepthEmu = depth;
            if (PptxShape3DCodec.TryReadContourWidth(shape3d, out var contourWidth))
                result.ShapeContourWidthEmu = contourWidth;
            if (PptxShape3DCodec.TryReadPresetMaterial(shape3d, out var presetMaterial))
                result.ShapePresetMaterial = presetMaterial;
            if (PptxShape3DCodec.TryReadBevelTopWidth(shape3d, out var bevelTopWidth))
                result.Shape3DBevelTopWidthEmu = bevelTopWidth;
            if (PptxShape3DCodec.TryReadBevelTopHeight(shape3d, out var bevelTopHeight))
                result.Shape3DBevelTopHeightEmu = bevelTopHeight;
            if (PptxShape3DCodec.TryReadBevelTopPreset(shape3d, out var bevelTopPreset))
                result.Shape3DBevelTopPreset = bevelTopPreset;
            if (PptxShape3DCodec.TryReadBevelBottomWidth(shape3d, out var bevelBottomWidth))
                result.Shape3DBevelBottomWidthEmu = bevelBottomWidth;
            if (PptxShape3DCodec.TryReadBevelBottomHeight(shape3d, out var bevelBottomHeight))
                result.Shape3DBevelBottomHeightEmu = bevelBottomHeight;
            if (PptxShape3DCodec.TryReadBevelBottomPreset(shape3d, out var bevelBottomPreset))
                result.Shape3DBevelBottomPreset = bevelBottomPreset;
            if (PptxShape3DCodec.TryReadContourRgb(shape3d, out var contourRgb))
                result.Shape3DContourRgb = contourRgb;
            if (PptxShape3DCodec.TryReadExtrusionRgb(shape3d, out var extrusionRgb))
                result.Shape3DExtrusionRgb = extrusionRgb;
            if (PptxShape3DCodec.TryReadContourColorScheme(shape3d, out var contourColorScheme))
                result.Shape3DContourColorScheme = contourColorScheme;
            if (PptxShape3DCodec.TryReadExtrusionColorScheme(shape3d, out var extrusionColorScheme))
                result.Shape3DExtrusionColorScheme = extrusionColorScheme;
        }
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } scene3d &&
            PptxShape3DCodec.TryReadSceneCameraPreset(scene3d, out var sceneCameraPreset))
            result.Shape3DSceneCameraPreset = sceneCameraPreset;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } cameraZoomScene3d &&
            PptxShape3DCodec.TryReadSceneCameraZoom(cameraZoomScene3d, out var sceneCameraZoom))
            result.Shape3DSceneCameraZoomThousandthPercent = sceneCameraZoom;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } cameraFovScene3d &&
            PptxShape3DCodec.TryReadSceneCameraFov(cameraFovScene3d, out var sceneCameraFov))
            result.Shape3DSceneCameraFov60000 = sceneCameraFov;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } lightRigScene3d &&
            PptxShape3DCodec.TryReadSceneLightRigPreset(lightRigScene3d, out var sceneLightRigPreset))
            result.Shape3DSceneLightRigPreset = sceneLightRigPreset;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } lightRigDirectionScene3d &&
            PptxShape3DCodec.TryReadSceneLightRigDirection(lightRigDirectionScene3d, out var sceneLightRigDirection))
            result.Shape3DSceneLightRigDirection = sceneLightRigDirection;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationLatitudeScene3d &&
            PptxShape3DCodec.TryReadSceneLightRigRotationLatitude(lightRigRotationLatitudeScene3d, out var sceneLightRigRotationLatitude))
            result.Shape3DSceneLightRigRotationLatitude60000 = sceneLightRigRotationLatitude;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationLongitudeScene3d &&
            PptxShape3DCodec.TryReadSceneLightRigRotationLongitude(lightRigRotationLongitudeScene3d, out var sceneLightRigRotationLongitude))
            result.Shape3DSceneLightRigRotationLongitude60000 = sceneLightRigRotationLongitude;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } lightRigRotationRevolutionScene3d &&
            PptxShape3DCodec.TryReadSceneLightRigRotationRevolution(lightRigRotationRevolutionScene3d, out var sceneLightRigRotationRevolution))
            result.Shape3DSceneLightRigRotationRevolution60000 = sceneLightRigRotationRevolution;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationLatitudeScene3d &&
            PptxShape3DCodec.TryReadSceneCameraRotationLatitude(cameraRotationLatitudeScene3d, out var sceneCameraRotationLatitude))
            result.Shape3DSceneCameraRotationLatitude60000 = sceneCameraRotationLatitude;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationLongitudeScene3d &&
            PptxShape3DCodec.TryReadSceneCameraRotationLongitude(cameraRotationLongitudeScene3d, out var sceneCameraRotationLongitude))
            result.Shape3DSceneCameraRotationLongitude60000 = sceneCameraRotationLongitude;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } cameraRotationRevolutionScene3d &&
            PptxShape3DCodec.TryReadSceneCameraRotationRevolution(cameraRotationRevolutionScene3d, out var sceneCameraRotationRevolution))
            result.Shape3DSceneCameraRotationRevolution60000 = sceneCameraRotationRevolution;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropAnchorX(backdropAnchorScene3d, out var sceneBackdropAnchorX))
            result.Shape3DSceneBackdropAnchorXEmu = sceneBackdropAnchorX;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorYScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropAnchorY(backdropAnchorYScene3d, out var sceneBackdropAnchorY))
            result.Shape3DSceneBackdropAnchorYEmu = sceneBackdropAnchorY;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropAnchorZScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropAnchorZ(backdropAnchorZScene3d, out var sceneBackdropAnchorZ))
            result.Shape3DSceneBackdropAnchorZEmu = sceneBackdropAnchorZ;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDxScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropNormalDx(backdropNormalDxScene3d, out var sceneBackdropNormalDx))
            result.Shape3DSceneBackdropNormalDxEmu = sceneBackdropNormalDx;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDyScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropNormalDy(backdropNormalDyScene3d, out var sceneBackdropNormalDy))
            result.Shape3DSceneBackdropNormalDyEmu = sceneBackdropNormalDy;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropNormalDzScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropNormalDz(backdropNormalDzScene3d, out var sceneBackdropNormalDz))
            result.Shape3DSceneBackdropNormalDzEmu = sceneBackdropNormalDz;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDxScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropUpDx(backdropUpDxScene3d, out var sceneBackdropUpDx))
            result.Shape3DSceneBackdropUpDxEmu = sceneBackdropUpDx;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDyScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropUpDy(backdropUpDyScene3d, out var sceneBackdropUpDy))
            result.Shape3DSceneBackdropUpDyEmu = sceneBackdropUpDy;
        if (properties?.Elements<A.Scene3DType>().Count() == 1 &&
            properties.GetFirstChild<A.Scene3DType>() is { } backdropUpDzScene3d &&
            PptxShape3DCodec.TryReadSceneBackdropUpDz(backdropUpDzScene3d, out var sceneBackdropUpDz))
            result.Shape3DSceneBackdropUpDzEmu = sceneBackdropUpDz;
        result.Accessibility = PptxNonVisualAccessibilityCodec.Read(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);
        return result;
    }

    private static bool IsSimpleShape(P.Shape shape, PptxPartContext slideContext, bool allowNegativeOffset = false)
    {
        if (shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is not null) return false;
        if (shape.ShapeStyle is not null) return false;
        var properties = shape.ShapeProperties;
        var transform = properties?.Transform2D;
        var geometry = Geometry(shape);
        if (properties is null || properties.Elements<A.Transform2D>().Count() != 1 ||
            !PptxShapeTransformCodec.Supports(
                transform,
                allowSingleZeroExtent: geometry == "line",
                allowNegativeOffset: allowNegativeOffset)) return false;
        if (geometry != "custom" &&
            !PptxPresetGeometryAdjustmentCodec.TryRead(properties.GetFirstChild<A.PresetGeometry>(), geometry, out _)) return false;
        var imageFill = properties.GetFirstChild<A.BlipFill>();
        var imageFillSupported = imageFill is not null && PptxShapeImageFillCodec.TryRead(imageFill, slideContext, out _);
        var linePathSupported = geometry == "custom" &&
            PpjLinePathCodec.IsLineLike(ReadShape(shape, slideContext));
        if (geometry == "custom")
        {
            var frame = ReadFrame(shape);
            var customGeometrySupported = PptxCustomGeometryCodec.Supports(properties.GetFirstChild<A.CustomGeometry>(), frame.Width, frame.Height);
            // Some third-party decks use a legal custom path profile that the
            // semantic geometry reader does not model (for example, a
            // rectangle-only path list without an adjustment list). When
            // that shape is source-bound to a strict image fill, its native
            // geometry can be preserved verbatim while the owning frame is
            // edited. Do not widen this exception to solid-filled or
            // source-free shapes: those still need a fully modeled path
            // graph before they become editable.
            if (!customGeometrySupported && !imageFillSupported) return false;
        }
        // A source-bound custom geometry may carry an explicit empty
        // alphaModFix on its image fill.  That effect is a no-op, but the
        // ordinary image-paint projection intentionally remains strict so
        // authored shapes and regular pictures do not silently widen their
        // semantic surface.  The shape-level image-fill profile is the
        // bounded path for this imported custom-geometry case.
        if (!SimpleFill(properties, slideContext,
                allowSourceBoundCustomImageFill: geometry == "custom" && imageFillSupported)) return false;
        var outline = properties.GetFirstChild<A.Outline>();
        if (!PptxLineStyleCodec.TryRead(outline, out var lineStyle)) return false;
        if (!string.Equals(geometry, "line", StringComparison.Ordinal) && !linePathSupported &&
            (lineStyle.StartArrow.Length > 0 || lineStyle.EndArrow.Length > 0)) return false;
        if (!PptxShadowCodec.TryRead(properties, out _) &&
            !PptxGlowCodec.TryRead(properties, out _) &&
            !PptxInnerShadowCodec.TryRead(properties, out _) &&
            !PptxReflectionCodec.TryRead(properties, out _) &&
            !PptxSoftEdgeCodec.TryRead(properties, out _)) return false;
        if (properties.ChildElements.Any(child => child is not A.Transform2D and not A.PresetGeometry and not A.CustomGeometry and not A.NoFill and not A.SolidFill and not A.GradientFill and not A.BlipFill and not A.Outline and not A.EffectList)) return false;
        // A literal single stroked custom path is the source-bound line
        // profile.  It has no text body by design, but its path/stroke/frame
        // leaves are still safely editable through setLinePath/setStroke.
        return PptxTextCodec.SupportsEditing(shape.TextBody) ||
            geometry == "custom" && imageFillSupported && shape.TextBody is null ||
            linePathSupported;
    }

    private static bool SimpleFill(
        OpenXmlCompositeElement element,
        PptxPartContext slideContext,
        bool allowSourceBoundCustomImageFill = false)
    {
        var fills = element.ChildElements.Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill).ToArray();
        if (fills.Length > 1) return false;
        if (fills.Length == 0 || fills[0] is A.NoFill) return true;
        if (fills[0] is A.GradientFill gradient) return PptxGradientFillCodec.TryRead(gradient, out _);
        if (fills[0] is A.BlipFill image)
            return PptxImagePaintCodec.TryRead(image, slideContext, out _) ||
                allowSourceBoundCustomImageFill && PptxShapeImageFillCodec.TryRead(image, slideContext, out _);
        var solid = (A.SolidFill)fills[0];
        if (PptxColor.TryDirectSolidSchemeWithOpacity(solid, out _, out var schemeOpacity))
            return schemeOpacity is null;
        if (solid.ChildElements.Count != 1 || solid.FirstChild is not A.RgbColorModelHex color || !HasOnlyAttributes(color, "val")) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        return color.ChildElements.Count == alphas.Length && alphas.Length <= 1 &&
               (alphas.Length == 0 || alphas[0].Val?.Value is >= 0 and <= 100_000 && HasOnlyAttributes(alphas[0], "val"));
    }

    private static uint? ReadFillOpacity(A.SolidFill? solid)
    {
        if (PptxColor.TryDirectSolidRgbWithOpacity(solid, out _, out var rgbOpacity) && rgbOpacity is not null)
            return rgbOpacity;
        if (PptxColor.TryDirectSolidSchemeWithOpacity(solid, out _, out var schemeOpacity) && schemeOpacity is not null)
            return schemeOpacity;
        return null;
    }

    private static void ApplyShape(P.Shape shape, PresentationElement source, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        var sourceHasBackgroundFill = shape.UseBackgroundFill?.HasValue == true;
        if (sourceHasBackgroundFill != semantic.HasUseBackgroundFill ||
            sourceHasBackgroundFill && shape.UseBackgroundFill!.Value != semantic.UseBackgroundFill)
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation shape {source.Id} cannot change its source-bound useBgFill attribute.");
        var properties = shape.ShapeProperties ??= new P.ShapeProperties();
        var sourceHasImageFill = PptxShapeImageFillCodec.TryRead(properties.GetFirstChild<A.BlipFill>(), slideContext, out var sourceImageFill);
        if (sourceHasImageFill)
        {
            if (string.IsNullOrWhiteSpace(semantic.ImageFillAssetId) || !semantic.ImageFillAssetId.Equals(sourceImageFill.Id, StringComparison.Ordinal))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation shape {source.Id} cannot change its source-bound image fill asset.");
            if (!string.IsNullOrWhiteSpace(semantic.FillRgb) || semantic.HasFillOpacityThousandthPercent || !string.IsNullOrWhiteSpace(semantic.FillScheme) || semantic.GradientFill is not null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation shape {source.Id} cannot replace a source-bound image fill with a solid or theme fill.");
        }
        else if (!string.IsNullOrWhiteSpace(semantic.ImageFillAssetId))
        {
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation shape {source.Id} has an image-fill identity that no longer matches its source.");
        }
        var transform = properties.Transform2D ??= new A.Transform2D();
        var offset = transform.Offset ??= new A.Offset();
        offset.X = semantic.LeftEmu;
        offset.Y = semantic.TopEmu;
        var extents = transform.Extents ??= new A.Extents();
        extents.Cx = semantic.WidthEmu;
        extents.Cy = semantic.HeightEmu;
        PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        // Imported image-filled custom geometries are intentionally source-owned
        // when their path graph is outside the semantic projection. Preserve
        // that native geometry while allowing the owning frame to move. A
        // populated custom path list remains the normal editable geometry path.
        if (!sourceHasImageFill || semantic.CustomPaths.Count > 0)
        PptxCustomGeometryCodec.Apply(properties, semantic, source.Id);
        PptxNonVisualAccessibilityCodec.ApplyBound(shape.NonVisualShapeProperties?.NonVisualDrawingProperties, semantic.Accessibility);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } actionOwner)
        {
            PptxHyperlinkCodec.ApplyElementAction(actionOwner, semantic.Action, slideContext);
            PptxHyperlinkCodec.ApplyElementHoverAction(actionOwner, semantic.HoverAction, slideContext);
        }
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties)
            drawingProperties.TextBox = semantic.Geometry == "textbox" ? true : null;
        if ((!sourceHasImageFill || semantic.ImageFill is not null) && !FillMatches(properties, semantic, slideContext))
            ReplaceFill(properties, semantic, slideContext);
        PptxLineStyleCodec.Apply(properties, semantic);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
            nonVisual.Name = source.Name;
        if (semantic.Glow is not null)
            PptxGlowCodec.Apply(properties, semantic.Glow);
        else if (semantic.InnerShadow is not null)
            PptxInnerShadowCodec.Apply(properties, semantic.InnerShadow);
        else if (semantic.Reflection is not null)
            PptxReflectionCodec.Apply(properties, semantic.Reflection);
        else if (semantic.SoftEdge is not null)
            PptxSoftEdgeCodec.Apply(properties, semantic.SoftEdge);
        else if (semantic.Shadow is not null)
            PptxShadowCodec.Apply(properties, semantic.Shadow);
        else if (PptxGlowCodec.TryRead(properties, out var sourceGlow) && sourceGlow is not null)
            PptxGlowCodec.Apply(properties, null);
        else if (PptxInnerShadowCodec.TryRead(properties, out var sourceInnerShadow) && sourceInnerShadow is not null)
            PptxInnerShadowCodec.Apply(properties, null);
        else if (PptxReflectionCodec.TryRead(properties, out var sourceReflection) && sourceReflection is not null)
            PptxReflectionCodec.Apply(properties, null);
        else if (PptxSoftEdgeCodec.TryRead(properties, out var sourceSoftEdge) && sourceSoftEdge is not null)
            PptxSoftEdgeCodec.Apply(properties, null);
        else
            PptxShadowCodec.Apply(properties, null);
        PptxTextCodec.Apply(shape, semantic, slideContext);
    }

    private static CodecException UnsupportedPresentationEdit(int slideIndex, int elementIndex, OpenXmlPart slidePart) => new(
        "unsupported_presentation_edit",
        $"Presentation slide {slideIndex + 1} element {elementIndex + 1} is preserved but not safely editable by this codec slice.",
        PartPath(slidePart));

    private static void ValidateNativeObjectRequest(PresentationElement original, PresentationElement requested)
    {
        var allowed = original.Clone();
        allowed.Name = requested.Name;
        if (requested.HasHidden) allowed.Hidden = requested.Hidden;
        else allowed.ClearHidden();
        if (requested.HasLocked) allowed.Locked = requested.Locked;
        else allowed.ClearLocked();
        allowed.Opaque.LeftEmu = requested.Opaque.LeftEmu;
        allowed.Opaque.TopEmu = requested.Opaque.TopEmu;
        allowed.Opaque.WidthEmu = requested.Opaque.WidthEmu;
        allowed.Opaque.HeightEmu = requested.Opaque.HeightEmu;
        if (allowed.Opaque.OleWorkbook is not null && requested.Opaque.OleWorkbook is not null)
            allowed.Opaque.OleWorkbook.ReplacementAssetId = requested.Opaque.OleWorkbook.ReplacementAssetId;
        if (allowed.Opaque.OleOfficePackage is not null && requested.Opaque.OleOfficePackage is not null)
            allowed.Opaque.OleOfficePackage.ReplacementAssetId = requested.Opaque.OleOfficePackage.ReplacementAssetId;
        if (allowed.Opaque.DiagramText is not null && requested.Opaque.DiagramText is not null)
        {
            allowed.Opaque.DiagramText.Nodes.Clear();
            allowed.Opaque.DiagramText.Nodes.Add(requested.Opaque.DiagramText.Nodes);
        }
        // Source binding equality is checked against the actual source above;
        // reuse the caller's equivalent instance to keep protobuf equality
        // focused on the semantic payload.
        allowed.Source = requested.Source.Clone();
        if (!allowed.Equals(requested))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation native object {requested.Id} may edit only its name, outer frame, and an explicitly recognized OLE payload.");
    }

    private static bool NativePlacementChanged(PresentationElement original, PresentationElement requested) =>
        original.Name != requested.Name ||
        original.Opaque.LeftEmu != requested.Opaque.LeftEmu ||
        original.Opaque.TopEmu != requested.Opaque.TopEmu ||
        original.Opaque.WidthEmu != requested.Opaque.WidthEmu ||
        original.Opaque.HeightEmu != requested.Opaque.HeightEmu;

    private static void ApplyNativePlacement(OpenXmlElement source, PresentationElement requested)
    {
        var frame = requested.Opaque;
        if (source is P.Picture picture)
        {
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(picture.ShapeProperties?.GetFirstChild<A.Transform2D>() ??
                throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported picture placement owner."), frame);
            return;
        }
        if (source is P.ConnectionShape connector)
        {
            if (connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(connector.ShapeProperties?.GetFirstChild<A.Transform2D>() ??
                throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported connector placement owner."), frame);
            return;
        }
        if (source is P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(graphicFrame.Transform!, frame);
            if (PptxNativeObjectCatalog.Classify(source) == "oleObject")
            {
                // PowerPoint stores a second transform on the OLE preview
                // picture. Keep it derived from the outer frame so Office and
                // fallback renderers agree after a move/resize.
                var previewTransform = graphicFrame.Descendants<A.Transform2D>().FirstOrDefault();
                if (previewTransform is not null) SetFrame(previewTransform, frame);
            }
            return;
        }
        if (source is P.GroupShape group)
        {
            if (group.GetFirstChild<P.NonVisualGroupShapeProperties>()?.NonVisualDrawingProperties is { } nonVisual)
                nonVisual.Name = requested.Name;
            SetFrame(group.GetFirstChild<P.GroupShapeProperties>()!.GetFirstChild<A.TransformGroup>()!, frame);
            return;
        }
        throw new CodecException("unsupported_presentation_edit", $"Presentation native object {requested.Id} has no supported placement owner.");
    }

    private static void SetFrame(P.Transform transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.Transform2D transform, PresentationOpaqueElement frame)
    {
        transform.Offset ??= new A.Offset();
        transform.Extents ??= new A.Extents();
        transform.Offset.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static void SetFrame(A.TransformGroup transform, PresentationOpaqueElement frame)
    {
        transform.Offset!.X = frame.LeftEmu;
        transform.Offset.Y = frame.TopEmu;
        transform.Extents!.Cx = frame.WidthEmu;
        transform.Extents.Cy = frame.HeightEmu;
    }

    private static OpenXmlElement BuildFill(PresentationShape source, PptxPartContext context)
    {
        if (source.ImageFill is not null)
            return PptxImagePaintCodec.Build(source.ImageFill, context, "shape fill");
        if (source.GradientFill is not null)
            return PptxGradientFillCodec.Build(source.GradientFill, "Presentation shape fill");
        if (!string.IsNullOrWhiteSpace(source.FillScheme))
            return PptxColor.BuildSolidScheme(source.FillScheme, source.HasFillOpacityThousandthPercent ? source.FillOpacityThousandthPercent : null);
        if (string.IsNullOrWhiteSpace(source.FillRgb)) return new A.NoFill();
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(source.FillRgb) };
        if (source.HasFillOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)source.FillOpacityThousandthPercent) });
        return new A.SolidFill(color);
    }

    private static void ReplaceFill(OpenXmlCompositeElement parent, PresentationShape source, PptxPartContext context)
    {
        var currentImageRelationshipId = PptxImagePaintCodec.RelationshipId(parent.GetFirstChild<A.BlipFill>());
        foreach (var child in parent.ChildElements.Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill).ToArray()) child.Remove();
        var fill = BuildFill(source, context);
        var replacementImageRelationshipId = PptxImagePaintCodec.RelationshipId(fill as A.BlipFill);
        var reference = parent.ChildElements.FirstOrDefault(child => child is A.Outline || child.LocalName is "effectLst" or "effectDag" or "scene3d" or "sp3d");
        if (reference is null) parent.Append(fill);
        else parent.InsertBefore(fill, reference);
        context.RemoveIfUnreferenced(currentImageRelationshipId == replacementImageRelationshipId ? string.Empty : currentImageRelationshipId);
    }

    private static bool FillMatches(OpenXmlCompositeElement parent, PresentationShape source, PptxPartContext context)
    {
        if (source.ImageFill is not null)
            return PptxImagePaintCodec.TryRead(parent.GetFirstChild<A.BlipFill>(), context, out var current) &&
                current.Equals(source.ImageFill);
        if (source.GradientFill is not null)
            return PptxGradientFillCodec.TryRead(parent.GetFirstChild<A.GradientFill>(), out var current) &&
                current.Equals(source.GradientFill);
        var requested = string.IsNullOrWhiteSpace(source.FillRgb) ? string.Empty : PptxColor.Normalize(source.FillRgb);
        var opacity = source.HasFillOpacityThousandthPercent ? source.FillOpacityThousandthPercent : (uint?)null;
        if (parent.GetFirstChild<A.NoFill>() is not null) return requested.Length == 0 && opacity is null;
        var solidFill = parent.GetFirstChild<A.SolidFill>();
        var solid = PptxColor.SolidRgb(solidFill);
        if (solid.Length > 0)
            return requested.Equals(solid, StringComparison.OrdinalIgnoreCase) && ReadFillOpacity(solidFill) == opacity;
        var requestedScheme = string.IsNullOrWhiteSpace(source.FillScheme) ? string.Empty : PptxColor.NormalizeScheme(source.FillScheme);
        var scheme = PptxColor.SolidSchemeWithOpacity(solidFill);
        if (scheme.Length > 0)
            return requestedScheme.Equals(scheme, StringComparison.OrdinalIgnoreCase) && ReadFillOpacity(solidFill) == opacity;
        return requested.Length == 0 && opacity is null && !parent.ChildElements.Any(child => child.LocalName.EndsWith("Fill", StringComparison.Ordinal));
    }

    private static void BuildPresentation(
        PresentationDocument package,
        IPptxSourceFreeBuildPlan plan,
        PptxAssetCatalog assetCatalog,
        Action<int, PresentationSlide>? validateSlide = null)
    {
        var artifact = plan.Presentation;
        var authoredMasters = artifact.Masters.Count == 0
            ? [new PresentationMaster { Id = "__officekit/default-master", Name = "Office Clean Room" }]
            : artifact.Masters.ToArray();
        var canonicalMasterId = authoredMasters[0].Id;
        var presentationPart = package.AddPresentationPart();
        var masterEntries = new List<(PresentationMaster Master, SlideMasterPart Part, ThemePart Theme, P.SlideLayoutIdList LayoutIds)>();
        for (var masterIndex = 0; masterIndex < authoredMasters.Length; masterIndex++)
        {
            var master = authoredMasters[masterIndex];
            var masterPart = presentationPart.AddNewPart<SlideMasterPart>($"rIdMaster{masterIndex + 1}");
            var ownerThemePart = masterPart.AddNewPart<ThemePart>($"rIdTheme{masterIndex + 1}");
            ownerThemePart.Theme = BasicTheme(artifact.AuthoredTheme);
            var layoutIds = new P.SlideLayoutIdList();
            masterPart.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(BasicShapeTree())
                {
                    Name = string.IsNullOrWhiteSpace(master.Name) ? "Office Clean Room" : master.Name,
                },
                BasicColorMap(),
                layoutIds,
                new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
            masterEntries.Add((master, masterPart, ownerThemePart, layoutIds));
        }
        var themePart = masterEntries[0].Theme;
        var sourceLayouts = artifact.Layouts.ToList();
        PresentationLayout? fallbackLayout = null;
        if (sourceLayouts.Count == 0 || artifact.Slides.Any(slide => string.IsNullOrWhiteSpace(slide.LayoutId)))
        {
            var fallbackId = "__officekit/default-layout";
            while (sourceLayouts.Any(layout => layout.Id.Equals(fallbackId, StringComparison.Ordinal))) fallbackId += "-";
            fallbackLayout = new PresentationLayout
            {
                Id = fallbackId,
                Name = "Blank",
                MasterId = canonicalMasterId,
                Type = "blank",
            };
            sourceLayouts.Insert(0, fallbackLayout);
        }

        var layoutEntries = new List<(PresentationLayout Layout, SlideLayoutPart Part, int MasterIndex)>();
        for (var layoutIndex = 0; layoutIndex < sourceLayouts.Count; layoutIndex++)
        {
            var sourceLayout = sourceLayouts[layoutIndex];
            var masterIndex = string.IsNullOrWhiteSpace(sourceLayout.MasterId)
                ? 0
                : Array.FindIndex(authoredMasters, master => master.Id.Equals(sourceLayout.MasterId, StringComparison.Ordinal));
            if (masterIndex < 0)
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {sourceLayout.Id} references missing master {sourceLayout.MasterId}.");
            var owner = masterEntries[masterIndex];
            var localLayoutIndex = owner.LayoutIds.ChildElements.Count;
            var relationshipId = $"rIdLayout{localLayoutIndex + 1}";
            var layoutPart = owner.Part.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.AddPart(owner.Part, $"rIdMaster{masterIndex + 1}");
            var layoutRoot = new P.SlideLayout(
                new P.CommonSlideData(BasicShapeTree()) { Name = string.IsNullOrWhiteSpace(sourceLayout.Name) ? "Blank" : sourceLayout.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()))
            { Preserve = true };
            layoutRoot.SetAttribute(new OpenXmlAttribute("type", string.Empty, SourceFreeLayoutType(sourceLayout)));
            layoutPart.SlideLayout = layoutRoot;
            owner.LayoutIds.Append(new P.SlideLayoutId
            {
                Id = checked(2_147_483_649U + (uint)layoutIndex),
                RelationshipId = relationshipId,
            });
            layoutEntries.Add((sourceLayout, layoutPart, masterIndex));
        }
        var layoutPartById = layoutEntries.ToDictionary(entry => entry.Layout.Id, entry => entry.Part, StringComparer.Ordinal);
        var defaultLayoutPart = fallbackLayout is not null
            ? layoutPartById[fallbackLayout.Id]
            : layoutEntries[0].Part;

        var slideIdList = new P.SlideIdList();
        var slideParts = new SlidePart[artifact.Slides.Count];
        // Authored source-free plans materialize element graphs lazily. Keep
        // only the pages that carry modern comments so the source-free
        // comment validator can resolve drawing monikers without retaining
        // every materialized slide until package finalization.
        var sourceFreeModernCommentSlides = artifact.Slides.ToArray();
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            var source = artifact.Slides[slideIndex];
            var relationshipId = $"rIdSlide{slideIndex + 1}";
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            slideParts[slideIndex] = slidePart;
            var layoutPart = string.IsNullOrWhiteSpace(source.LayoutId)
                ? defaultLayoutPart
                : layoutPartById.GetValueOrDefault(source.LayoutId) ??
                  throw new CodecException("invalid_presentation_layout", $"Presentation slide {source.Id} references missing layout {source.LayoutId}.");
            slidePart.AddPart(layoutPart, "rIdLayout1");
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(BasicShapeTree()) { Name = source.Name },
                new P.ColorMapOverride(new A.MasterColorMapping()));
            PptxSlideVisibilityCodec.BuildSourceFree(slidePart.Slide, source);
            slideIdList.Append(new P.SlideId { Id = checked((uint)(256 + slideIndex)), RelationshipId = relationshipId });
        }
        var slideIdByPartPath = slideParts
            .Select((part, index) => (Path: PartPath(part), Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var slidePartById = slideParts
            .Select((part, index) => (Part: part, Id: artifact.Slides[index].Id))
            .ToDictionary(item => item.Id, item => item.Part, StringComparer.Ordinal);
        var customShowCatalog = PptxCustomShowCatalog.From(artifact.CustomShows);
        foreach (var (master, ownerPart, _, _) in masterEntries)
        {
            var masterContext = new PptxPartContext(ownerPart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
            PptxBackgroundCodec.Build(ownerPart.SlideMaster.CommonSlideData!, master.Background, masterContext);
            PptxMasterTextStylesCodec.Build(ownerPart.SlideMaster, master.TextStyles, masterContext);
            var masterShapeTree = ownerPart.SlideMaster.CommonSlideData!.ShapeTree!;
            foreach (var (placeholder, index) in master.Placeholders.Select((placeholder, index) => (placeholder, index)))
                masterShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), masterContext));
        }
        foreach (var (layout, layoutPart, _) in layoutEntries)
        {
            var layoutContext = new PptxPartContext(layoutPart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog);
            var layoutCommon = layoutPart.SlideLayout!.CommonSlideData!;
            PptxBackgroundCodec.Build(layoutCommon, layout.Background, layoutContext);
            var layoutShapeTree = layoutCommon.ShapeTree!;
            foreach (var (placeholder, index) in layout.Placeholders.Select((placeholder, index) => (placeholder, index)))
                layoutShapeTree.Append(PptxPlaceholderCodec.Build(placeholder, checked((uint)(index + 2)), layoutContext));
        }
        PresentationSlide? previousSlide = null;
        for (var slideIndex = 0; slideIndex < artifact.Slides.Count; slideIndex++)
        {
            // Keep a fully materialized prior slide only when the next page
            // actually needs it for Morph lowering/validation. Ordinary pages
            // can release their element graph after the slide part is written,
            // avoiding a second page-sized object graph at the build high-water
            // mark without changing the Morph contract.
            var source = plan.MaterializeSlide(
                slideIndex,
                plan.RequiresPreviousSlide(slideIndex) ? previousSlide : null);
            if (source.ModernComments.Count > 0)
                sourceFreeModernCommentSlides[slideIndex] = source;
            validateSlide?.Invoke(slideIndex, source);
            var slidePart = slideParts[slideIndex];
            var slideCommon = slidePart.Slide!.CommonSlideData!;
            var slideContext = new PptxPartContext(slidePart, slideIdByPartPath, slidePartById, assetCatalog, customShowCatalog, slideNumber: checked(slideIndex + 1));
            PptxBackgroundCodec.Build(slideCommon, source.Background, slideContext);
            PptxTransitionCodec.Build(slidePart.Slide!, source.Transition);
            var shapeTree = slideCommon.ShapeTree!;
            var flattenedElements = FlattenPresentationElements(source.Elements).ToArray();
            plan.RecordNativeBindings(slideIndex, source, flattenedElements);
            var nativeIdsByElementId = flattenedElements.Select((element, index) => (element.Id, NativeId: checked((uint)(index + 2))))
                .ToDictionary(item => item.Id, item => item.NativeId, StringComparer.Ordinal);
            foreach (var element in source.Elements)
                shapeTree.Append(BuildElement(element, nativeIdsByElementId, slideContext, slidePart, package));
            PptxTimingCodec.ValidateMorphContext(source, previousSlide);
            PptxTimingCodec.Build(slidePart.Slide!, source, nativeIdsByElementId);
            slidePart.Slide.Save();
            slidePart.UnloadRootElement();
            previousSlide = plan.RequiresPreviousSlide(slideIndex + 1) ? source : null;
        }
        var notesMasterRelationshipId = PptxSpeakerNotesCodec.BuildSourceFree(presentationPart, themePart, slideParts, artifact.Slides);
        var presentationRoot = new P.Presentation();
        var masterIdList = new P.SlideMasterIdList();
        for (var masterIndex = 0; masterIndex < masterEntries.Count; masterIndex++)
            masterIdList.Append(new P.SlideMasterId
            {
                Id = checked(2_147_483_648U + (uint)masterIndex),
                RelationshipId = $"rIdMaster{masterIndex + 1}",
            });
        presentationRoot.Append(masterIdList);
        if (notesMasterRelationshipId is not null)
            presentationRoot.Append(new P.NotesMasterIdList(new P.NotesMasterId { Id = notesMasterRelationshipId }));
        presentationRoot.Append(
            slideIdList,
            new P.SlideSize
            {
                Cx = checked((int)(artifact.SlideWidthEmu > 0 ? artifact.SlideWidthEmu : DefaultSlideWidthEmu)),
                Cy = checked((int)(artifact.SlideHeightEmu > 0 ? artifact.SlideHeightEmu : DefaultSlideHeightEmu)),
            },
            new P.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L },
            new P.DefaultTextStyle());
        if (artifact.CustomShows.Count > 0)
        {
            var relationshipIdBySlideId = artifact.Slides
                .Select((slide, index) => (slide.Id, RelationshipId: $"rIdSlide{index + 1}"))
                .ToDictionary(item => item.Id, item => item.RelationshipId, StringComparer.Ordinal);
            PptxCustomShowCodec.BuildSourceFree(presentationRoot, artifact, relationshipIdBySlideId);
        }
        if (artifact.Sections.Count > 0)
        {
            var nativeSlideIdByPublicId = artifact.Slides
                .Select((slide, index) => (slide.Id, NativeId: checked((uint)(256 + index))))
                .ToDictionary(item => item.Id, item => item.NativeId, StringComparer.Ordinal);
            PptxSectionCodec.BuildSourceFree(presentationRoot, artifact, nativeSlideIdByPublicId);
        }
        presentationPart.Presentation = presentationRoot;
        PptxLegacyCommentsCodec.BuildSourceFree(presentationPart, slideParts, artifact.Slides);
        PptxModernCommentsCodec.BuildSourceFree(
            presentationPart,
            slideIdList.Elements<P.SlideId>().ToArray(),
            slideParts,
            sourceFreeModernCommentSlides);
        foreach (var (_, _, ownerTheme, _) in masterEntries) ownerTheme.Theme.Save();
        foreach (var (_, layoutPart, _) in layoutEntries) layoutPart.SlideLayout!.Save();
        foreach (var (_, ownerPart, _, _) in masterEntries) ownerPart.SlideMaster.Save();
        presentationPart.Presentation.Save();
    }

    private static string SourceFreeLayoutType(PresentationLayout layout)
    {
        var type = string.IsNullOrWhiteSpace(layout.Type) ? "blank" : layout.Type;
        return type switch
        {
            "blank" or "title" or "titleOnly" or "obj" => type,
            _ => throw new CodecException(
                "unsupported_presentation_features",
                $"Source-free presentation layout {layout.Id} uses unsupported type {type}. Use blank, title, titleOnly, or obj."),
        };
    }

    private static void ValidateSourceFreeTextPlaceholder(PresentationPlaceholder placeholder, string ownerId)
    {
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, placeholder.Id);
        if (placeholder.DirectFrame is null)
            throw new CodecException("invalid_presentation_placeholder", $"Source-free {ownerId} placeholder {placeholder.Id} requires a direct frame.");
    }

    private static void ValidateSourceFreeTextPlaceholderIdentity(PresentationPlaceholderIdentity placeholder, string elementId) =>
        ValidateSourceFreeTextPlaceholderType(placeholder.Type, elementId);

    private static void ValidateSourceFreeTextPlaceholderType(string type, string ownerId)
    {
        if (type is "title" or "body" or "ctrTitle" or "subTitle") return;
        throw new CodecException(
            "unsupported_presentation_features",
            $"Source-free presentation placeholder {ownerId} uses {type}; only title, body, ctrTitle, and subTitle text placeholders are supported.");
    }

    // Removal is an edit even when the imported list was empty or unmodeled.
    // Keep this intent out of semantic hashes, but still run the native writer
    // so it can remove a modeled empty list or refuse an unknown source list.
    private static bool HasParagraphTabStopRemovalIntent(PresentationElement element) => element.ContentCase switch
    {
        PresentationElement.ContentOneofCase.Shape => element.Shape.TextBody?.Paragraphs.Any(paragraph => paragraph.HasNoTabStops && paragraph.NoTabStops) == true,
        PresentationElement.ContentOneofCase.Group => element.Group.Children.Any(HasParagraphTabStopRemovalIntent),
        _ => false,
    };

    // A right-margin removal also requires native validation when projection
    // omitted the source attribute. Do not accept unknown state as a no-op.
    private static bool HasParagraphRightMarginRemovalIntent(PresentationElement element) => element.ContentCase switch
    {
        PresentationElement.ContentOneofCase.Shape => element.Shape.TextBody?.Paragraphs.Any(paragraph =>
            paragraph.RightMarginCase == PresentationTextParagraph.RightMarginOneofCase.NoMarginRight) == true,
        PresentationElement.ContentOneofCase.Group => element.Group.Children.Any(HasParagraphRightMarginRemovalIntent),
        _ => false,
    };

    private static bool ContainsAutomaticFields(PresentationElement element) => element.ContentCase switch
    {
        PresentationElement.ContentOneofCase.Shape => ContainsAutomaticFields(element.Shape.TextBody),
        PresentationElement.ContentOneofCase.Table => element.Table.Rows.Any(row => row.Cells.Any(cell => ContainsAutomaticFields(cell.TextBody))),
        PresentationElement.ContentOneofCase.Group => element.Group.Children.Any(ContainsAutomaticFields),
        PresentationElement.ContentOneofCase.Chart => ContainsAutomaticFields(element.Chart.TitleBody),
        _ => false,
    };

    private static bool ContainsAutomaticFields(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Any(paragraph => paragraph.Runs.Any(run =>
            run.ContentCase == PresentationTextRun.ContentOneofCase.Field && run.Field.Automatic));

    private static IEnumerable<PresentationElement> FlattenPresentationElements(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in FlattenPresentationElements(element.Group.Children)) yield return child;
        }
    }

    private static (PresentationElement[] Retained, PresentationElement[] Authored) SplitSourceBoundElements(
        PresentationSlide slide,
        int sourceElementCount,
        int slideIndex,
        OpenXmlPart owner)
    {
        var retainedCount = 0;
        while (retainedCount < slide.Elements.Count && slide.Elements[retainedCount].Source is not null) retainedCount++;
        var retained = slide.Elements.Take(retainedCount).ToArray();
        var authored = slide.Elements.Skip(retainedCount).ToArray();
        if (authored.Any(element => element.Source is not null))
            throw new CodecException(
                "presentation_element_topology_changed",
                $"Presentation slide {slideIndex + 1} source-bound elements must remain an ordered prefix before authored overlay elements.",
                PartPath(owner));
        if (retained.Length + slide.ElementDeletions.Count != sourceElementCount)
            throw new CodecException(
                "presentation_element_topology_changed",
                $"Source-preserving PPTX export requires slide {slideIndex + 1}'s original {sourceElementCount}-element topology to be covered by retained elements plus explicit deletions; the artifact contains {retained.Length} retained elements, {slide.ElementDeletions.Count} deletions, and {authored.Length} authored overlay elements.",
                PartPath(owner));
        foreach (var element in authored) ValidateAuthoredOverlayElement(element, slideIndex, owner);
        return (retained, authored);
    }

    private static void ValidateAuthoredOverlayElement(PresentationElement element, int slideIndex, OpenXmlPart owner)
    {
        if (BoundedAuthoredOverlayViolation(element) is { } violation)
            throw new CodecException(
                "unsupported_presentation_authored_overlay",
                $"Presentation slide {slideIndex + 1} authored overlay {element.Id} {violation}.",
                PartPath(owner));
    }

    internal static string? BoundedAuthoredOverlayViolation(PresentationElement element)
    {
        if (element.ContentCase == PresentationElement.ContentOneofCase.Image)
            return null;
        if (element.ContentCase != PresentationElement.ContentOneofCase.Shape)
            return "must be a canonical textbox, basic shape, or embedded rectangular image";
        var shape = element.Shape;
        if (!string.IsNullOrEmpty(shape.CatalogIconName))
        {
            if (!PpjIconCatalog.Contains(shape.CatalogIconName))
                return "uses an icon name outside the pinned compiler catalog";
            if (shape.Geometry != "custom" || shape.CustomPaths.Count != 1 ||
                shape.Placeholder is not null || shape.DirectFrame is not null || shape.HasUseBackgroundFill ||
                shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
                shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 ||
                shape.TextRectangle is not null || shape.TextBody is not null || shape.ImageFill is not null ||
                !string.IsNullOrEmpty(shape.ImageFillAssetId))
                return "uses state outside the bounded named-icon overlay profile";
            try
            {
                PptxCustomGeometryCodec.Validate(shape, element.Id);
            }
            catch (CodecException)
            {
                return "contains invalid named-icon custom geometry";
            }
            return null;
        }
        if (shape.Geometry is not ("textbox" or "rect" or "roundRect" or "ellipse") ||
            shape.Placeholder is not null || shape.DirectFrame is not null || shape.HasUseBackgroundFill ||
            shape.CustomPaths.Count > 0 || shape.CustomAdjustments.Count > 0 || shape.CustomGuides.Count > 0 ||
            shape.CustomConnectionSites.Count > 0 || shape.CustomAdjustmentHandles.Count > 0 || shape.TextRectangle is not null)
            return "uses geometry or layout identity outside the bounded textbox/basic-shape profile";
        var paragraphs = (shape.TextBody?.Paragraphs ?? []).Concat(shape.TextBody?.ListStyles ?? []);
        return paragraphs.Any(paragraph =>
            paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet ||
            paragraph.Runs.Any(run => run.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.RunHyperlink))
                ? "cannot add picture or hyperlink relationships"
                : null;
    }

    private static IReadOnlyDictionary<string, uint> AuthoredOverlayNativeIds(
        IReadOnlyList<OpenXmlElement> sourceElements,
        IReadOnlyList<PresentationElement> authored,
        int slideIndex,
        OpenXmlPart owner)
    {
        var occupied = sourceElements.SelectMany(PptxElementDeletionCodec.NativeIds).ToHashSet();
        var next = occupied.Count == 0 ? 1U : occupied.Max();
        var output = new Dictionary<string, uint>(StringComparer.Ordinal);
        foreach (var element in authored)
        {
            if (next == uint.MaxValue)
                throw new CodecException(
                    "unsupported_presentation_authored_overlay",
                    $"Presentation slide {slideIndex + 1} has no remaining native drawing ID for authored overlay {element.Id}.",
                    PartPath(owner));
            next++;
            if (!output.TryAdd(element.Id, next))
                throw new CodecException(
                    "invalid_presentation_element",
                    $"Presentation slide {slideIndex + 1} contains duplicate authored overlay identity {element.Id}.",
                    PartPath(owner));
        }
        return output;
    }

    private static OpenXmlElement BuildElement(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart,
        PresentationDocument? package = null)
    {
        OpenXmlElement output = element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => BuildShape(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Image => PptxPictureCodec.Build(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Table => PptxTableCodec.Build(element, nativeIdsByElementId[element.Id], slideContext),
            PresentationElement.ContentOneofCase.Connector => PptxConnectorCodec.Build(element, nativeIdsByElementId[element.Id], nativeIdsByElementId),
            PresentationElement.ContentOneofCase.Chart => PptxChartCodec.Build(element, nativeIdsByElementId[element.Id], slidePart, slideContext),
            PresentationElement.ContentOneofCase.Diagram => PptxSmartArtCodec.Build(element, nativeIdsByElementId[element.Id], slideContext, slidePart),
            PresentationElement.ContentOneofCase.Group => BuildGroup(element, nativeIdsByElementId, slideContext, slidePart, package),
            PresentationElement.ContentOneofCase.Media when package is not null =>
                PptxMediaCodec.Build(element, nativeIdsByElementId[element.Id], slideContext, slidePart, package),
            PresentationElement.ContentOneofCase.Media => throw new CodecException(
                "unsupported_presentation_authored_overlay",
                $"Presentation media {element.Id} can be authored only in a source-free PPJ build."),
            _ => throw new CodecException("unsupported_presentation_element", $"Opaque presentation element {element.Id} requires its validated source package and cannot be authored from scratch."),
        };
        PptxElementStateCodec.ApplyAuthored(output, element);
        return output;
    }

    private static P.GroupShape BuildGroup(
        PresentationElement element,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        PptxPartContext slideContext,
        SlidePart slidePart,
        PresentationDocument? package)
    {
        var group = element.Group;
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeIdsByElementId[element.Id], Name = element.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, group.Accessibility);
        var transform = new A.TransformGroup(
            new A.Offset { X = group.LeftEmu, Y = group.TopEmu },
            new A.Extents { Cx = group.WidthEmu, Cy = group.HeightEmu },
            new A.ChildOffset { X = group.ChildLeftEmu, Y = group.ChildTopEmu },
            new A.ChildExtents { Cx = group.ChildWidthEmu, Cy = group.ChildHeightEmu });
        PptxFrameTransformCodec.Apply(transform, group.FrameTransform);
        var output = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                nonVisual,
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(transform));
        foreach (var child in group.Children)
            output.Append(BuildElement(child, nativeIdsByElementId, slideContext, slidePart, package));
        return output;
    }

    private static P.Shape BuildShape(PresentationElement source, uint nativeId, PptxPartContext slideContext)
    {
        var semantic = source.Shape;
        if (!string.IsNullOrWhiteSpace(semantic.ImageFillAssetId))
            throw new CodecException(
                "unsupported_presentation_features",
                $"Presentation shape {source.Id} cannot author a source-bound image fill outside an imported source package.");
        var directFrame = semantic.DirectFrame;
        var transform = new A.Transform2D(
            new A.Offset { X = directFrame?.LeftEmu ?? semantic.LeftEmu, Y = directFrame?.TopEmu ?? semantic.TopEmu },
            new A.Extents { Cx = directFrame?.WidthEmu ?? semantic.WidthEmu, Cy = directFrame?.HeightEmu ?? semantic.HeightEmu });
        if (directFrame is not null)
        {
            transform.Rotation = directFrame.HasRotationAngle60000 ? directFrame.RotationAngle60000 : null;
            transform.HorizontalFlip = directFrame.HasFlipHorizontal ? directFrame.FlipHorizontal : null;
            transform.VerticalFlip = directFrame.HasFlipVertical ? directFrame.FlipVertical : null;
        }
        else
        {
            PptxShapeTransformCodec.Apply(transform, semantic.Transform);
        }
        var properties = new P.ShapeProperties(transform);
        PptxCustomGeometryCodec.Apply(properties, semantic, source.Id);
        properties.Append(BuildFill(semantic, slideContext));
        properties.Append(PptxLineStyleCodec.Build(semantic));
        PptxShadowCodec.Apply(properties, semantic.Shadow);
        PptxGlowCodec.Apply(properties, semantic.Glow);
        PptxInnerShadowCodec.Apply(properties, semantic.InnerShadow);
        PptxReflectionCodec.Apply(properties, semantic.Reflection);
        PptxSoftEdgeCodec.Apply(properties, semantic.SoftEdge);
        var applicationProperties = new P.ApplicationNonVisualDrawingProperties();
        if (semantic.Placeholder is not null)
        {
            var nativePlaceholder = new P.PlaceholderShape { Index = semantic.Placeholder.Index };
            nativePlaceholder.SetAttribute(new OpenXmlAttribute("type", string.Empty, semantic.Placeholder.Type));
            applicationProperties.Append(nativePlaceholder);
        }
        var nonVisual = new P.NonVisualDrawingProperties { Id = nativeId, Name = source.Name };
        PptxNonVisualAccessibilityCodec.ApplyAuthored(nonVisual, semantic.Accessibility);
        if (semantic.Action is not null)
            PptxHyperlinkCodec.AppendElementAction(nonVisual, semantic.Action, slideContext);
        if (semantic.HoverAction is not null)
            PptxHyperlinkCodec.AppendElementHoverAction(nonVisual, semantic.HoverAction, slideContext);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                nonVisual,
                new P.NonVisualShapeDrawingProperties { TextBox = semantic.Geometry == "textbox" ? true : null },
                applicationProperties),
            properties,
            PptxTextCodec.Build(semantic, slideContext));
    }

    private static bool ApplyGroup(
        P.GroupShape source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location,
        EffectiveCodecLimits limits)
    {
        if (original.ContentCase != PresentationElement.ContentOneofCase.Group || requested.ContentCase != PresentationElement.ContentOneofCase.Group)
            throw new CodecException("presentation_group_content_changed", $"Presentation slide {slideIndex + 1} {location} changed its group content type.", PartPath(slideContext.Owner));
        var sourceChildren = GroupElements(source);
        if (sourceChildren.Length != original.Group.Children.Count || sourceChildren.Length != requested.Group.Children.Count)
            throw new CodecException("presentation_group_topology_changed", $"Presentation slide {slideIndex + 1} {location} changed its fixed group topology.", PartPath(slideContext.Owner));
        var zOrderPlan = AnalyzeElementZOrder(sourceChildren);
        var requestedSourceOrder = new int[requested.Group.Children.Count];
        var requestedSourceIndexes = new HashSet<int>();
        for (var requestedIndex = 0; requestedIndex < requested.Group.Children.Count; requestedIndex++)
        {
            var requestedChild = requested.Group.Children[requestedIndex];
            var binding = requestedChild.Source ?? throw new CodecException(
                "missing_presentation_element_binding",
                $"Presentation slide {slideIndex + 1} {location} child {requestedIndex + 1} is missing its source binding.",
                PartPath(slideContext.Owner));
            if (binding.ShapeTreeIndex >= (uint)sourceChildren.Length || !requestedSourceIndexes.Add((int)binding.ShapeTreeIndex))
                throw new CodecException(
                    "presentation_group_topology_changed",
                    $"Presentation slide {slideIndex + 1} {location} children must identify a unique local source order.",
                    PartPath(slideContext.Owner));
            requestedSourceOrder[requestedIndex] = (int)binding.ShapeTreeIndex;
        }
        var sourceOrderChanged = !requestedSourceOrder.SequenceEqual(Enumerable.Range(0, sourceChildren.Length));
        if (sourceOrderChanged && !zOrderPlan.Supported)
            throw new CodecException(
                "unsupported_presentation_element_reorder",
                $"Presentation slide {slideIndex + 1} {location} cannot safely reorder its group children: {zOrderPlan.BlockedReason}.",
                PartPath(slideContext.Owner));

        var changed = !Equals(original.Group.Accessibility, requested.Group.Accessibility);
        PptxNonVisualAccessibilityCodec.ApplyBound(
            source.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            requested.Group.Accessibility,
            "group");
        if (requested.Name != original.Name ||
            requested.Group.LeftEmu != original.Group.LeftEmu || requested.Group.TopEmu != original.Group.TopEmu ||
            requested.Group.WidthEmu != original.Group.WidthEmu || requested.Group.HeightEmu != original.Group.HeightEmu ||
            requested.Group.ChildLeftEmu != original.Group.ChildLeftEmu || requested.Group.ChildTopEmu != original.Group.ChildTopEmu ||
            requested.Group.ChildWidthEmu != original.Group.ChildWidthEmu || requested.Group.ChildHeightEmu != original.Group.ChildHeightEmu ||
            !Equals(requested.Group.FrameTransform, original.Group.FrameTransform))
        {
            source.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!.Name = requested.Name;
            var transform = source.GroupShapeProperties!.GetFirstChild<A.TransformGroup>()!;
            transform.Offset!.X = requested.Group.LeftEmu;
            transform.Offset.Y = requested.Group.TopEmu;
            transform.Extents!.Cx = requested.Group.WidthEmu;
            transform.Extents.Cy = requested.Group.HeightEmu;
            transform.ChildOffset!.X = requested.Group.ChildLeftEmu;
            transform.ChildOffset.Y = requested.Group.ChildTopEmu;
            transform.ChildExtents!.Cx = requested.Group.ChildWidthEmu;
            transform.ChildExtents.Cy = requested.Group.ChildHeightEmu;
            PptxFrameTransformCodec.Apply(transform, requested.Group.FrameTransform);
            changed = true;
        }
        for (var requestedIndex = 0; requestedIndex < sourceChildren.Length; requestedIndex++)
        {
            var sourceIndex = requestedSourceOrder[requestedIndex];
            var sourceChild = sourceChildren[sourceIndex];
            var originalChild = original.Group.Children[sourceIndex];
            var requestedChild = requested.Group.Children[requestedIndex];
            var binding = requestedChild.Source ?? throw new CodecException(
                "missing_presentation_element_binding",
                $"Presentation slide {slideIndex + 1} {location} child {requestedIndex + 1} is missing its source binding.",
                PartPath(slideContext.Owner));
            if (requestedChild.Id != originalChild.Id || binding.ShapeTreeIndex != (uint)sourceIndex ||
                !binding.ElementSha256.Equals(HashElement(sourceChild), StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != originalChild.Source?.Editable ||
                binding.TextEditable != originalChild.Source?.TextEditable ||
                binding.AccessibilityEditable != originalChild.Source?.AccessibilityEditable ||
                binding.VisibilityEditable != originalChild.Source?.VisibilityEditable ||
                binding.LockingEditable != originalChild.Source?.LockingEditable ||
                binding.ZOrderCapability is null ||
                binding.ZOrderCapability.Supported != zOrderPlan.Supported ||
                !binding.ZOrderCapability.BlockedReason.Equals(zOrderPlan.BlockedReason, StringComparison.Ordinal) ||
                !binding.SemanticSha256.Equals(originalChild.Source?.SemanticSha256 ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
                !SemanticHash(originalChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_element_binding_mismatch",
                    $"Presentation slide {slideIndex + 1} {location} child {requestedIndex + 1} does not match its owner-local source binding.",
                    PartPath(slideContext.Owner));
            if (!HasParagraphTabStopRemovalIntent(requestedChild) && !HasParagraphRightMarginRemovalIntent(requestedChild) && SemanticHash(requestedChild).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            var stateChanged = PptxElementStateCodec.StateChanged(originalChild, requestedChild);
            var contentChanged = !PptxElementStateCodec.EqualExceptState(originalChild, requestedChild);
            if (stateChanged)
            {
                PptxElementStateCodec.ApplyBound(sourceChild, originalChild, requestedChild);
                changed = true;
            }
            if (!contentChanged) continue;
            if (IsMediaAccessibilityOnlyChange(sourceChild, originalChild, requestedChild))
            {
                ApplyMediaAccessibility(sourceChild, requestedChild);
                changed = true;
                continue;
            }
            if (!binding.Editable)
            {
                if (binding.TextEditable &&
                    sourceChild is P.Shape sourceTextShape &&
                    IsBoundTextOnlyShapeEdit(originalChild, requestedChild, sourceTextShape))
                {
                    PptxTextCodec.Apply(sourceTextShape, requestedChild.Shape, slideContext);
                    changed = true;
                    continue;
                }
                throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} child {requestedIndex + 1} is read-only.", PartPath(slideContext.Owner));
            }
            changed |= ApplyGroupChild(sourceChild, originalChild, requestedChild, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, $"{location} child {requestedIndex + 1}", limits);
        }
        if (sourceOrderChanged)
        {
            ApplyGroupElementZOrder(source, sourceChildren, requestedSourceOrder);
            changed = true;
        }
        return changed;
    }

    private static bool ApplyGroupChild(
        OpenXmlElement source,
        PresentationElement original,
        PresentationElement requested,
        PptxPartContext slideContext,
        IReadOnlyDictionary<uint, string> elementIdsByNativeId,
        IReadOnlyDictionary<string, uint> nativeIdsByElementId,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes,
        int slideIndex,
        string location,
        EffectiveCodecLimits limits)
    {
        if (source is P.Shape shape && requested.ContentCase == PresentationElement.ContentOneofCase.Shape && IsSimpleShape(shape, slideContext, allowNegativeOffset: Geometry(shape) == "line"))
            ApplyShape(shape, requested, slideContext);
        else if (source is P.Picture picture && requested.ContentCase == PresentationElement.ContentOneofCase.Image && PptxPictureCodec.TryRead(picture, slideContext, out _))
            PptxPictureCodec.Apply(picture, requested, slideContext);
        else if (source is P.GraphicFrame table && requested.ContentCase == PresentationElement.ContentOneofCase.Table && PptxTableCodec.TryRead(table, slideContext, out _))
            PptxTableCodec.Apply(table, requested, slideContext);
        else if (source is P.ConnectionShape connector && requested.ContentCase == PresentationElement.ContentOneofCase.Connector && PptxConnectorCodec.TryRead(connector, elementIdsByNativeId, out _))
            PptxConnectorCodec.Apply(connector, requested, nativeIdsByElementId);
        else if (source is P.GraphicFrame chart && requested.ContentCase == PresentationElement.ContentOneofCase.Chart && PptxChartCodec.TryRead(chart, slideContext, out _, out var chartEditable) && chartEditable)
        {
            var replacement = PptxChartCodec.Apply(chart, requested, slideContext, limits);
            changedParts.Add(replacement.PartPath);
            changedParts.UnionWith(replacement.ChangedPartPaths);
            replacedOpaquePartHashes.Add(replacement.PartPath, replacement.Sha256);
            foreach (var entry in replacement.ReplacedPartHashes) replacedOpaquePartHashes.Add(entry.Key, entry.Value);
            return replacement.SlideChanged;
        }
        else if (source is P.GroupShape group && requested.ContentCase == PresentationElement.ContentOneofCase.Group && original.ContentCase == PresentationElement.ContentOneofCase.Group && TryReadGroup(group, original.Id, slideContext, elementIdsByNativeId, out _))
            return ApplyGroup(group, original, requested, slideContext, elementIdsByNativeId, nativeIdsByElementId, changedParts, replacedOpaquePartHashes, slideIndex, location, limits);
        else if (requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
                 PptxNativeObjectCatalog.SupportsPlacementEditing(source))
        {
            // A nested opaque child may be moved or resized only through the
            // same direct-frame proof as a top-level native object.  Do not
            // silently drop an OLE/diagram replacement or any other semantic
            // mutation because this path has no asset catalog for it.
            ValidateNativePlacementRequest(original, requested);
            if (NativePlacementChanged(original, requested))
                ApplyNativePlacement(source, requested);
        }
        else
            throw new CodecException("unsupported_presentation_edit", $"Presentation slide {slideIndex + 1} {location} changed outside the bounded group-child profile.", PartPath(slideContext.Owner));
        return true;
    }

    private static void ValidateNativePlacementRequest(PresentationElement original, PresentationElement requested)
    {
        var allowed = original.Clone();
        allowed.Name = requested.Name;
        if (requested.HasHidden) allowed.Hidden = requested.Hidden;
        else allowed.ClearHidden();
        if (requested.HasLocked) allowed.Locked = requested.Locked;
        else allowed.ClearLocked();
        allowed.Opaque.LeftEmu = requested.Opaque.LeftEmu;
        allowed.Opaque.TopEmu = requested.Opaque.TopEmu;
        allowed.Opaque.WidthEmu = requested.Opaque.WidthEmu;
        allowed.Opaque.HeightEmu = requested.Opaque.HeightEmu;
        // Source binding is validated by ApplyGroup before this helper. Keep
        // the comparison focused on the bounded native payload instead of
        // rejecting an equivalent protobuf instance with a different object
        // identity.
        allowed.Source = requested.Source.Clone();
        if (!allowed.Equals(requested))
            throw new CodecException(
                "unsupported_presentation_edit",
                $"Presentation native object {requested.Id} may edit only its name and outer frame inside an imported group.");
    }

    private static bool IsMediaAccessibilityOnlyChange(
        OpenXmlElement source,
        PresentationElement original,
        PresentationElement requested) =>
        original.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
        requested.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
        source is P.Picture &&
        PptxNativeObjectCatalog.IsMediaPicture(source) &&
        !Equals(original.Opaque.Accessibility, requested.Opaque.Accessibility) &&
        AccessibilityOnlyRequest(original, requested);

    private static bool AccessibilityOnlyRequest(PresentationElement original, PresentationElement requested)
    {
        var expected = original.Clone();
        expected.Opaque.Accessibility = requested.Opaque.Accessibility?.Clone();
        return expected.Equals(requested);
    }

    private static void ApplyMediaAccessibility(OpenXmlElement source, PresentationElement requested)
    {
        if (source is not P.Picture picture)
            throw new CodecException("unsupported_presentation_edit", "Source media accessibility requires a picture-shaped media owner.");
        PptxNonVisualAccessibilityCodec.ApplyResidualBound(
            picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
            requested.Opaque.Accessibility,
            "media");
    }

    private static P.ShapeTree BasicShapeTree() => new(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L },
            new A.ChildOffset { X = 0L, Y = 0L },
            new A.ChildExtents { Cx = 0L, Cy = 0L })));

    private static P.ColorMap BasicColorMap() => new()
    {
        Background1 = A.ColorSchemeIndexValues.Light1,
        Text1 = A.ColorSchemeIndexValues.Dark1,
        Background2 = A.ColorSchemeIndexValues.Light2,
        Text2 = A.ColorSchemeIndexValues.Dark2,
        Accent1 = A.ColorSchemeIndexValues.Accent1,
        Accent2 = A.ColorSchemeIndexValues.Accent2,
        Accent3 = A.ColorSchemeIndexValues.Accent3,
        Accent4 = A.ColorSchemeIndexValues.Accent4,
        Accent5 = A.ColorSchemeIndexValues.Accent5,
        Accent6 = A.ColorSchemeIndexValues.Accent6,
        Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
        FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
    };

    private static A.Theme BasicTheme(PresentationThemeArtifact? authored = null)
    {
        var defaults = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        var accents = defaults
            .Select((value, index) => ThemeRgb(authored is not null && index < authored.AccentRgb.Count
                ? authored.AccentRgb[index]
                : value))
            .ToArray();
        var accentTransforms = authored?.AccentTransforms
            .ToDictionary(transform => transform.Role, StringComparer.Ordinal)
            ?? new Dictionary<string, PresentationThemeColorTransform>(StringComparer.Ordinal);
        var majorFont = authored?.HasMajorFontFamily == true ? authored.MajorFontFamily : "Arial";
        var majorFontEastAsia = authored?.HasMajorFontFamilyEastAsia == true ? authored.MajorFontFamilyEastAsia : majorFont;
        var majorFontComplexScript = authored?.HasMajorFontFamilyComplexScript == true ? authored.MajorFontFamilyComplexScript : majorFont;
        var minorFont = authored?.HasMinorFontFamily == true ? authored.MinorFontFamily : majorFont;
        var minorFontEastAsia = authored?.HasMinorFontFamilyEastAsia == true ? authored.MinorFontFamilyEastAsia : minorFont;
        var minorFontComplexScript = authored?.HasMinorFontFamilyComplexScript == true ? authored.MinorFontFamilyComplexScript : minorFont;
        var themeName = authored?.HasName == true ? authored.Name : "Office Clean Room";
        (string Rgb, uint? Opacity)? dark1 = authored?.HasDark1Rgb == true ? ThemeRgb(authored.Dark1Rgb) : null;
        (string Rgb, uint? Opacity)? light1 = authored?.HasLight1Rgb == true ? ThemeRgb(authored.Light1Rgb) : null;
        var dark2 = authored?.HasDark2Rgb == true ? ThemeRgb(authored.Dark2Rgb) : ThemeRgb("1F497D");
        var light2 = authored?.HasLight2Rgb == true ? ThemeRgb(authored.Light2Rgb) : ThemeRgb("EEECE1");
        var hyperlink = authored?.HasHyperlinkRgb == true ? ThemeRgb(authored.HyperlinkRgb) : ThemeRgb("0000FF");
        var followedHyperlink = authored?.HasFollowedHyperlinkRgb == true ? ThemeRgb(authored.FollowedHyperlinkRgb) : ThemeRgb("800080");
        return new A.Theme(
        new A.ThemeElements(
                new A.ColorScheme(
                new A.Dark1Color(dark1 is null
                    ? new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }
                    : ThemeRgb(dark1.Value)),
                new A.Light1Color(light1 is null
                    ? new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }
                    : ThemeRgb(light1.Value)),
                new A.Dark2Color(ThemeRgb(dark2)),
                new A.Light2Color(ThemeRgb(light2)),
                new A.Accent1Color(ThemeRgb(accents[0], accentTransforms.GetValueOrDefault("accent1"))),
                new A.Accent2Color(ThemeRgb(accents[1], accentTransforms.GetValueOrDefault("accent2"))),
                new A.Accent3Color(ThemeRgb(accents[2], accentTransforms.GetValueOrDefault("accent3"))),
                new A.Accent4Color(ThemeRgb(accents[3], accentTransforms.GetValueOrDefault("accent4"))),
                new A.Accent5Color(ThemeRgb(accents[4], accentTransforms.GetValueOrDefault("accent5"))),
                new A.Accent6Color(ThemeRgb(accents[5], accentTransforms.GetValueOrDefault("accent6"))),
                new A.Hyperlink(ThemeRgb(hyperlink)),
                new A.FollowedHyperlinkColor(ThemeRgb(followedHyperlink))) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = majorFont }, new A.EastAsianFont { Typeface = majorFontEastAsia }, new A.ComplexScriptFont { Typeface = majorFontComplexScript }),
                new A.MinorFont(new A.LatinFont { Typeface = minorFont }, new A.EastAsianFont { Typeface = minorFontEastAsia }, new A.ComplexScriptFont { Typeface = minorFontComplexScript })) { Name = themeName },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                new A.LineStyleList(
                    BasicThemeOutline(9_525),
                    BasicThemeOutline(25_400),
                    BasicThemeOutline(38_100)),
                new A.EffectStyleList(
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList()),
                    new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))) { Name = themeName }))
        { Name = themeName };
    }

    private static A.Outline BasicThemeOutline(int width) => new(
        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
        new A.PresetDash { Val = A.PresetLineDashValues.Solid })
    { Width = width, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    private static (string Rgb, uint? Opacity) ThemeRgb(string value)
    {
        var rgb = PptxColor.NormalizeThemeRgb(value, out var opacity);
        return (rgb, opacity);
    }

    private static A.RgbColorModelHex ThemeRgb((string Rgb, uint? Opacity) value)
    {
        var color = new A.RgbColorModelHex { Val = value.Rgb };
        if (value.Opacity is { } opacity) color.Append(new A.Alpha { Val = checked((int)opacity) });
        return color;
    }

    private static A.RgbColorModelHex ThemeRgb(
        (string Rgb, uint? Opacity) value,
        PresentationThemeColorTransform? transform)
    {
        var color = new A.RgbColorModelHex { Val = value.Rgb };
        if (transform?.HasTintThousandth == true)
            color.Append(new A.Tint { Val = checked((int)transform.TintThousandth) });
        if (transform?.HasShadeThousandth == true)
            color.Append(new A.Shade { Val = checked((int)transform.ShadeThousandth) });
        if (transform?.HasLuminanceModulationThousandth == true)
            color.Append(new A.LuminanceModulation { Val = checked((int)transform.LuminanceModulationThousandth) });
        if (transform?.HasLuminanceOffsetThousandth == true)
            color.Append(new A.LuminanceOffset { Val = checked((int)transform.LuminanceOffsetThousandth) });
        if (transform?.HasSaturationModulationThousandth == true)
            color.Append(new A.SaturationModulation { Val = checked((int)transform.SaturationModulationThousandth) });
        if (transform?.HasSaturationOffsetThousandth == true)
            color.Append(new A.SaturationOffset { Val = checked((int)transform.SaturationOffsetThousandth) });
        if (transform?.HasRedModulationThousandth == true)
            color.Append(new A.RedModulation { Val = checked((int)transform.RedModulationThousandth) });
        if (transform?.HasRedOffsetThousandth == true)
            color.Append(new A.RedOffset { Val = checked(transform.RedOffsetThousandth) });
        if (transform?.HasGreenModulationThousandth == true)
            color.Append(new A.GreenModulation { Val = checked((int)transform.GreenModulationThousandth) });
        if (transform?.HasGreenOffsetThousandth == true)
            color.Append(new A.GreenOffset { Val = checked(transform.GreenOffsetThousandth) });
        if (transform?.HasBlueModulationThousandth == true)
            color.Append(new A.BlueModulation { Val = checked((int)transform.BlueModulationThousandth) });
        if (transform?.HasBlueOffsetThousandth == true)
            color.Append(new A.BlueOffset { Val = checked(transform.BlueOffsetThousandth) });
        if (transform?.HasHueModulationThousandth == true)
            color.Append(new A.HueModulation { Val = checked((int)transform.HueModulationThousandth) });
        if (transform?.HasHueOffsetAngleThousandth == true)
            color.Append(new A.HueOffset { Val = checked(transform.HueOffsetAngleThousandth) });
        if (value.Opacity is { } opacity)
            color.Append(new A.Alpha { Val = checked((int)opacity) });
        if (transform?.HasAlphaModulationThousandth == true)
            color.Append(new A.AlphaModulation { Val = checked((int)transform.AlphaModulationThousandth) });
        if (transform?.HasAlphaOffsetThousandth == true)
            color.Append(new A.AlphaOffset { Val = checked((int)transform.AlphaOffsetThousandth) });
        if (transform?.HasGray == true)
            color.Append(new A.Gray());
        if (transform?.HasComp == true)
            color.Append(new A.Complement());
        if (transform?.HasInv == true)
            color.Append(new A.Inverse());
        if (transform?.HasGamma == true)
            color.Append(new A.Gamma());
        if (transform?.HasInvGamma == true)
            color.Append(new A.InverseGamma());
        return color;
    }

    private static OpenXmlElement[] ShapeElements(P.ShapeTree shapeTree) =>
        shapeTree.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static OpenXmlElement[] GroupElements(P.GroupShape group) =>
        group.ChildElements.Where(child => child is not P.NonVisualGroupShapeProperties and not P.GroupShapeProperties).ToArray();

    private static IReadOnlyDictionary<uint, string> NativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId,
        List<PptxNativeBinding>? previewBindings = null, string? partPath = null)
    {
        var output = new Dictionary<uint, string>();
        CollectNativeElementIds(elements, ownerId, output, previewBindings, ownerId, partPath);
        return output;
    }

    private static void CollectNativeElementIds(IReadOnlyList<OpenXmlElement> elements, string ownerId, IDictionary<uint, string> output,
        List<PptxNativeBinding>? previewBindings, string pageId, string? partPath)
    {
        for (var index = 0; index < elements.Count; index++)
        {
            var elementId = $"{ownerId}/element/{index + 1}";
            var nativeId = elements[index].Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ??
                           elements[index].Descendants<P14.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value;
            if (nativeId is not null)
            {
                output[nativeId.Value] = elementId;
                // Unlike the timing lookup, retain duplicates in provenance so
                // a later join can reject ambiguous native identities.
                previewBindings?.Add(new(pageId, elementId, elements[index].LocalName, partPath!, nativeId.Value));
            }
            if (elements[index] is P.GroupShape group)
                CollectNativeElementIds(GroupElements(group), elementId, output, previewBindings, pageId, partPath);
        }
    }

    private static (long Left, long Top, long Width, long Height) ReadFrame(OpenXmlElement element)
    {
        if (element is P.GraphicFrame graphicFrame && graphicFrame.Transform?.Offset is { } graphicOffset && graphicFrame.Transform.Extents is { } graphicExtents)
            return (graphicOffset.X?.Value ?? 0, graphicOffset.Y?.Value ?? 0, graphicExtents.Cx?.Value ?? 0, graphicExtents.Cy?.Value ?? 0);
        if (element is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { Offset: { } groupOffset, Extents: { } groupExtents })
            return (groupOffset.X?.Value ?? 0, groupOffset.Y?.Value ?? 0, groupExtents.Cx?.Value ?? 0, groupExtents.Cy?.Value ?? 0);
        if (element is P.ContentPart contentPart && contentPart.Transform2D is { Offset: { } contentOffset, Extents: { } contentExtents })
            return (contentOffset.X?.Value ?? 0, contentOffset.Y?.Value ?? 0, contentExtents.Cx?.Value ?? 0, contentExtents.Cy?.Value ?? 0);
        var transform = element.Descendants<A.Transform2D>().FirstOrDefault();
        if (transform?.Offset is not null && transform.Extents is not null)
            return (transform.Offset.X?.Value ?? 0, transform.Offset.Y?.Value ?? 0, transform.Extents.Cx?.Value ?? 0, transform.Extents.Cy?.Value ?? 0);
        var offset = element.Descendants<A.Offset>().FirstOrDefault();
        var extents = element.Descendants<A.Extents>().FirstOrDefault();
        return (offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, extents?.Cx?.Value ?? 0, extents?.Cy?.Value ?? 0);
    }

    private static string Geometry(P.Shape shape)
    {
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.TextBox?.Value == true) return "textbox";
        if (shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>() is not null) return "custom";
        var native = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>();
        var value = native?.Preset?.Value;
        if (value is null) return "rect";
        if (PptxCustomGeometryCodec.TryPresetName(value.Value, out var name)) return name;
        var raw = native!.GetAttribute("prst", string.Empty).Value;
        return string.IsNullOrEmpty(raw) ? "rect" : raw;
    }

    private static string ElementName(OpenXmlElement element, int index) =>
        element.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ??
        element.Descendants<P14.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value ??
        $"{element.LocalName} {index + 1}";

    private static string DescendantText(OpenXmlElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants<A.Text>().Select(text => text.Text));

    private static string SemanticHash(PresentationElement element)
    {
        var semantic = element.Clone();
        var placementEditable = semantic.ContentCase == PresentationElement.ContentOneofCase.Opaque && semantic.Source?.Editable == true;
        ClearElementIdentity(semantic);
        NormalizeSemanticForHash(semantic);
        if (placementEditable) semantic.Opaque.RawXml = string.Empty;
        return Hash(semantic.ToByteArray());
    }

    private static void NormalizeSemanticForHash(PresentationElement element)
    {
        // Imported visibility/locking readers materialize an explicit false
        // for an absent native attribute.  Treat that default presence as
        // equivalent to omission so a content-only source edit does not fail
        // post-write semantic validation.
        if (element.HasHidden && !element.Hidden) element.ClearHidden();
        if (element.HasLocked && !element.Locked) element.ClearLocked();
        if (element.ContentCase == PresentationElement.ContentOneofCase.Shape)
        {
            // catalog_icon_name is compiler provenance used only to admit the
            // bounded authored overlay. It is intentionally absent from PPTX
            // and therefore cannot participate in native semantic equality.
            element.Shape.CatalogIconName = string.Empty;
            PptxTextCodec.NormalizeSemantics(element.Shape);
            PptxLineStyleCodec.NormalizeSemantics(element.Shape);
            element.Shape.FillRgb = string.IsNullOrWhiteSpace(element.Shape.FillRgb) ? string.Empty : PptxColor.Normalize(element.Shape.FillRgb);
            element.Shape.FillScheme = string.IsNullOrWhiteSpace(element.Shape.FillScheme) ? string.Empty : PptxColor.NormalizeScheme(element.Shape.FillScheme);
            return;
        }
        if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
        {
            foreach (var child in element.Group.Children) NormalizeSemanticForHash(child);
            return;
        }
        if (element.ContentCase == PresentationElement.ContentOneofCase.Table)
        {
            // Normalize edit markers and the same uniform text/body choice
            // used by the table reader, including removal of the last bodyPr.
            foreach (var cell in element.Table.Rows.SelectMany(row => row.Cells))
                PptxTableCodec.NormalizeTextSemantics(cell);
            return;
        }
        if (element.ContentCase == PresentationElement.ContentOneofCase.Chart)
        {
            // Optional chart settings retain presence: removing explicit false
            // must still reach the ChartPart writer.
            if (element.Chart.TitleBody is not null)
            {
                var title = new PresentationShape
                {
                    Text = element.Chart.Title,
                    TextBody = element.Chart.TitleBody.Clone(),
                };
                PptxTextCodec.NormalizeSemantics(title);
                element.Chart.TitleBody = title.TextBody;
            }
            NormalizeChartAxisForHash(element.Chart.XAxis, "bottom");
            NormalizeChartAxisForHash(element.Chart.YAxis, "left");
            NormalizeChartAxisForHash(element.Chart.SecondaryXAxis, "top");
            NormalizeChartAxisForHash(element.Chart.SecondaryYAxis, "right");
        }
    }

    private static void NormalizeChartAxisForHash(SpreadsheetChartAxisArtifact? axis, string defaultPosition)
    {
        // c:axPos is mandatory in ChartML, but its canonical primary and
        // secondary values are represented by omission in PPJ. Normalize an
        // explicit request for the matching default before source-hash
        // comparison so authored and projected forms stay equivalent.
        if (axis?.HasPosition == true && string.Equals(axis.Position, defaultPosition, StringComparison.Ordinal))
            axis.ClearPosition();

        // tickLabelPosition=none is the native representation of hidden tick
        // labels. Older PPJ projections also carried tickLabelsVisible=false
        // as a compatibility alias, so canonicalize both forms before the
        // post-write semantic comparison.
        if (axis?.HasTickLabelsVisible == true && !axis.TickLabelsVisible &&
            (!axis.HasTickLabelPosition || string.Equals(axis.TickLabelPosition, "none", StringComparison.Ordinal)))
        {
            axis.TickLabelPosition = "none";
            axis.ClearTickLabelsVisible();
        }

        // OOXML's native default is to show tick labels. Treat an explicit
        // request for that default as equivalent to an omitted tickLblPos.
        if (axis?.HasTickLabelsVisible == true && axis.TickLabelsVisible)
            axis.ClearTickLabelsVisible();

        // A direct c:spPr line implies visibility. The native reader records
        // that implication explicitly, while the authored/source-bound
        // semantic request may omit the redundant presence bit.
        if (axis?.MajorGridlineStyle is not null && axis.HasMajorGridlineVisible && axis.MajorGridlineVisible)
            axis.ClearMajorGridlineVisible();
        if (axis?.MinorGridlineStyle is not null && axis.HasMinorGridlineVisible && axis.MinorGridlineVisible)
            axis.ClearMinorGridlineVisible();
    }

    private static void ClearElementIdentity(PresentationElement element)
    {
        // A source-bound opaque leaf with a proven placement profile may
        // change its native transform without changing its opaque payload.
        // Groups use this recursive semantic hash, so scrub that leaf's raw
        // XML before clearing identity just as the top-level hash does.
        if (element.ContentCase == PresentationElement.ContentOneofCase.Opaque && element.Source?.Editable == true)
            element.Opaque.RawXml = string.Empty;
        element.Id = string.Empty;
        element.Source = null;
        if (element.ContentCase != PresentationElement.ContentOneofCase.Group) return;
        foreach (var child in element.Group.Children) ClearElementIdentity(child);
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string DataPartPath(DataPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string RelationshipPartPath(OpenXmlPart part)
    {
        var path = PartPath(part);
        var separator = path.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : path[..separator];
        var fileName = separator < 0 ? path : path[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }
    private static byte[] PartBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
    private static byte[] DataPartBytes(DataPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
    private static void CopyPartBytes(OpenXmlPart source, OpenXmlPart target)
    {
        using var input = source.GetStream(FileMode.Open, FileAccess.Read);
        using var output = target.GetStream(FileMode.Create, FileAccess.Write);
        input.CopyTo(output);
    }
    private static void CopyDataPartBytes(DataPart source, DataPart target)
    {
        using var input = source.GetStream(FileMode.Open, FileAccess.Read);
        target.FeedData(input);
    }
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static PptxAssetCatalog ValidateEnvelope(
        ArtifactEnvelope envelope,
        EffectiveCodecLimits limits,
        bool allowEmptyExistingPictureAssets = false)
    {
        var context = ValidateEnvelopeHeader(envelope, limits, allowEmptyExistingPictureAssets);
        for (var slideIndex = 0; slideIndex < envelope.Presentation.Slides.Count; slideIndex++)
            ValidatePresentationSlide(envelope.Presentation.Slides[slideIndex], slideIndex, context, limits);
        return context.AssetCatalog;
    }

    private static EnvelopeValidationContext ValidateEnvelopeHeader(
        ArtifactEnvelope envelope,
        EffectiveCodecLimits limits,
        bool allowEmptyExistingPictureAssets = false)
    {
        if (envelope.ProtocolVersion != CodecWireProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Presentation || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Presentation)
            throw new CodecException("invalid_presentation_artifact", "Artifact envelope does not contain a presentation payload.");
        if (envelope.Presentation.Slides.Count == 0)
            throw new CodecException("missing_slides", "Presentation must contain at least one slide.");
        if ((uint)envelope.Presentation.Slides.Count > limits.MaxSheets)
            throw new CodecException("slide_budget_exceeded", $"Presentation has {envelope.Presentation.Slides.Count} slides and exceeds max_sheets ({limits.MaxSheets}).");
        if (envelope.Presentation.SlideWidthEmu < 0 || envelope.Presentation.SlideHeightEmu < 0 || envelope.Presentation.SlideWidthEmu > int.MaxValue || envelope.Presentation.SlideHeightEmu > int.MaxValue)
            throw new CodecException("invalid_slide_size", "Presentation slide dimensions must fit the PresentationML signed 32-bit EMU range.");
        var assetCatalog = new PptxAssetCatalog(
            envelope.Assets,
            limits,
            allowEmptyExistingPictureAssets: allowEmptyExistingPictureAssets);
        var hasSourcePackage = envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false };
        PptxViewPropertiesCodec.Validate(envelope.Presentation.ViewProperties, hasSourcePackage);

        if (envelope.Presentation.Masters.Count > 64)
            throw new CodecException("presentation_master_budget_exceeded", "Presentation cannot contain more than 64 slide masters.");
        if ((uint)envelope.Presentation.Layouts.Count > limits.MaxSheets)
            throw new CodecException("presentation_layout_budget_exceeded", $"Presentation has {envelope.Presentation.Layouts.Count} layouts and exceeds max_sheets ({limits.MaxSheets}).");
        ulong items = 0;
        PptxCustomShowCodec.Validate(envelope.Presentation, hasSourcePackage, limits, ref items);
        PptxSectionCodec.Validate(envelope.Presentation, hasSourcePackage, limits, ref items);
        var masterIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var master in envelope.Presentation.Masters)
        {
            if (string.IsNullOrWhiteSpace(master.Id) || !masterIds.Add(master.Id))
                throw new CodecException("invalid_presentation_master", "Presentation master IDs must be non-empty and unique.");
            if (master.Name.Length > 1_024)
                throw new CodecException("invalid_presentation_master", $"Presentation master {master.Id} name exceeds 1024 characters.");
            PptxMasterTextStylesCodec.Validate(master.TextStyles);
            PptxBackgroundCodec.Validate(master.Background, assetCatalog);
            ValidatePlaceholders(master.Id, master.Placeholders, assetCatalog, limits, ref items);
            foreach (var paragraph in MasterStyleParagraphs(master.TextStyles))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        var layoutIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var layout in envelope.Presentation.Layouts)
        {
            if (string.IsNullOrWhiteSpace(layout.Id) || !layoutIds.Add(layout.Id))
                throw new CodecException("invalid_presentation_layout", "Presentation layout IDs must be non-empty and unique.");
            if (!masterIds.Contains(layout.MasterId))
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} references missing master {layout.MasterId}.");
            if (layout.Name.Length > 1_024 || layout.Type.Length > 128)
                throw new CodecException("invalid_presentation_layout", $"Presentation layout {layout.Id} has invalid name or type metadata.");
            PptxBackgroundCodec.Validate(layout.Background, assetCatalog);
            ValidatePlaceholders(layout.Id, layout.Placeholders, assetCatalog, limits, ref items);
        }

        var mastersById = envelope.Presentation.Masters.ToDictionary(master => master.Id, StringComparer.Ordinal);
        var layoutsById = envelope.Presentation.Layouts.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
        if (!hasSourcePackage)
        {
            foreach (var master in envelope.Presentation.Masters)
                foreach (var placeholder in master.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, master.Id);
            foreach (var layout in envelope.Presentation.Layouts)
            {
                _ = SourceFreeLayoutType(layout);
                foreach (var placeholder in layout.Placeholders)
                    ValidateSourceFreeTextPlaceholder(placeholder, layout.Id);
            }
        }

        return new EnvelopeValidationContext(assetCatalog, hasSourcePackage, layoutIds, mastersById, layoutsById, items);
    }

    private static void ValidatePresentationSlide(
        PresentationSlide slide,
        int slideIndex,
        EnvelopeValidationContext context,
        EffectiveCodecLimits limits)
    {
        var assetCatalog = context.AssetCatalog;
        var hasSourcePackage = context.HasSourcePackage;
        var layoutIds = context.LayoutIds;
        var mastersById = context.MastersById;
        var layoutsById = context.LayoutsById;
        var items = context.Items;
        PptxSpeakerNotesCodec.Validate(slide.SpeakerNotes);
        PptxBackgroundCodec.Validate(slide.Background, assetCatalog);
        PptxTransitionCodec.Validate(slide.Transition);
        PptxLegacyCommentsCodec.Validate(slide, slideIndex);
        PptxModernCommentsCodec.Validate(slide, slideIndex, hasSourcePackage);
        if (!string.IsNullOrWhiteSpace(slide.LayoutId) && !layoutIds.Contains(slide.LayoutId))
            throw new CodecException("invalid_presentation_layout", $"Presentation slide {slide.Id} references missing layout {slide.LayoutId}.");
        if (!hasSourcePackage)
        {
            var placeholderShapes = slide.Elements.Where(element =>
                element.ContentCase == PresentationElement.ContentOneofCase.Shape && element.Shape.Placeholder is not null).ToArray();
            if (placeholderShapes.Length > 0)
            {
                if (string.IsNullOrWhiteSpace(slide.LayoutId) || !layoutsById.TryGetValue(slide.LayoutId, out var layout))
                    throw new CodecException("invalid_presentation_layout", $"Source-free presentation slide {slide.Id} has placeholders but no explicit layout binding.");
                var master = mastersById[layout.MasterId];
                foreach (var element in placeholderShapes)
                {
                    var placeholder = element.Shape.Placeholder;
                    if (placeholder.InheritsGeometry || element.Shape.DirectFrame is null)
                        throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
                    ValidateSourceFreeTextPlaceholderIdentity(placeholder, element.Id);
                    if (!master.Placeholders.Concat(layout.Placeholders).Any(candidate =>
                            candidate.Type.Equals(placeholder.Type, StringComparison.Ordinal) && candidate.Index == placeholder.Index))
                        throw new CodecException("presentation_placeholder_binding_mismatch", $"Source-free presentation slide placeholder {element.Id} has no matching master/layout placeholder.");
                }
            }
        }
        foreach (var element in slide.Elements)
            ValidatePresentationElement(element, hasSourcePackage, assetCatalog, limits, ref items, 0);
        context.Items = items;
    }

    private static void ValidatePresentationElement(
        PresentationElement element,
        bool hasSourcePackage,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items,
        int depth,
        bool allowNegativeOffset = false)
    {
        items++;
        if (items > limits.MaxCells)
            throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        if (depth > 16)
            throw new CodecException("presentation_group_depth_exceeded", "Presentation groups cannot be nested more than 16 levels.");
        if (string.IsNullOrWhiteSpace(element.Id) || element.Id.Length > 1_024 || element.Name.Length > 1_024)
            throw new CodecException("invalid_presentation_element", "Presentation element IDs and names must be bounded non-empty metadata.");

        // The source-preserving path re-reads this exact bound element before
        // applying any request, verifies its native and semantic hashes, and
        // rejects every semantic change when Editable is false. Treat that
        // contract as one opaque branch here rather than re-implementing a
        // partial validator for every DrawingML feature we deliberately do
        // not model. The envelope-wide item/depth/id bounds above still apply.
        if (hasSourcePackage && element.Source?.Editable == false) return;

        if (element.ContentCase == PresentationElement.ContentOneofCase.Shape)
        {
            if (element.Shape.HasUseBackgroundFill && !hasSourcePackage)
                throw new CodecException("unsupported_presentation_features", $"Presentation shape {element.Id} cannot author useBgFill without a validated source package.");
            if (element.Shape.Placeholder is not null && !hasSourcePackage &&
                (element.Shape.Placeholder.InheritsGeometry || element.Shape.DirectFrame is null))
                throw new CodecException("invalid_presentation_placeholder", $"Source-free presentation slide placeholder {element.Id} must use a direct frame.");
            if (element.Shape.Placeholder is not null && element.Shape.Transform is not null)
                throw new CodecException("invalid_presentation_transform", $"Presentation placeholder shape {element.Id} cannot carry an ordinary shape transform.");
            var inheritedPlaceholderGeometry = element.Shape.Placeholder?.InheritsGeometry == true &&
                element.Shape.DirectFrame is null && element.Source?.Editable == false;
            var freeLine = element.Shape.Geometry == "line";
            var sourceBoundFreeLine = hasSourcePackage && element.Source?.Editable == true && freeLine;
            var invalidExtent = freeLine
                ? element.Shape.WidthEmu == 0 && element.Shape.HeightEmu == 0
                : element.Shape.WidthEmu == 0 || element.Shape.HeightEmu == 0;
            if ((!inheritedPlaceholderGeometry && (!allowNegativeOffset && !sourceBoundFreeLine && (element.Shape.LeftEmu < 0 || element.Shape.TopEmu < 0) ||
                    element.Shape.WidthEmu < 0 || element.Shape.HeightEmu < 0 ||
                    invalidExtent)) ||
                element.Shape.LineWidthEmu < 0 || element.Shape.LineWidthEmu > int.MaxValue)
                throw new CodecException(
                    "invalid_presentation_frame",
                    $"Presentation shape {element.Id} has an invalid {element.Shape.Geometry} frame " +
                    $"({element.Shape.LeftEmu},{element.Shape.TopEmu},{element.Shape.WidthEmu},{element.Shape.HeightEmu}).");
            if (element.Shape.DirectFrame is not null)
            {
                if (element.Shape.Placeholder is null || element.Shape.Placeholder.InheritsGeometry)
                    throw new CodecException("invalid_presentation_placeholder", $"Presentation shape {element.Id} has inconsistent direct placeholder geometry.");
                PptxPlaceholderCodec.ValidateDirectFrame(element.Shape.DirectFrame, element.Id);
            }
            if (freeLine && element.Shape.Placeholder is not null)
                throw new CodecException("unsupported_presentation_geometry", $"Presentation free line {element.Id} cannot be a placeholder.");
            PptxCustomGeometryCodec.Validate(element.Shape, element.Id);
            if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb)) PptxColor.Normalize(element.Shape.FillRgb);
            if (!string.IsNullOrWhiteSpace(element.Shape.FillScheme) && !hasSourcePackage)
                throw new CodecException("unsupported_presentation_features", $"Source-free presentation shape {element.Id} cannot author a theme fill directly.");
            if (!string.IsNullOrWhiteSpace(element.Shape.ImageFillAssetId))
            {
                if (!hasSourcePackage)
                    throw new CodecException("unsupported_presentation_features", $"Source-free presentation shape {element.Id} cannot author an image fill.");
                _ = assetCatalog.Get(element.Shape.ImageFillAssetId);
            }
            if (element.Shape.HasFillOpacityThousandthPercent &&
                ((string.IsNullOrWhiteSpace(element.Shape.FillRgb) && string.IsNullOrWhiteSpace(element.Shape.FillScheme)) || element.Shape.GradientFill is not null ||
                 element.Shape.ImageFill is not null ||
                 element.Shape.FillOpacityThousandthPercent > 100_000))
                throw new CodecException("invalid_presentation_fill", $"Presentation shape {element.Id} has invalid solid-fill opacity.");
            if (element.Shape.GradientFill is not null)
            {
                if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb) || !string.IsNullOrWhiteSpace(element.Shape.FillScheme) || element.Shape.ImageFill is not null)
                    throw new CodecException("invalid_presentation_fill", $"Presentation shape {element.Id} cannot combine gradient and solid fill state.");
                PptxGradientFillCodec.Validate(element.Shape.GradientFill, $"Presentation shape {element.Id}");
            }
            if (element.Shape.ImageFill is not null)
            {
                if (!string.IsNullOrWhiteSpace(element.Shape.FillRgb) || !string.IsNullOrWhiteSpace(element.Shape.FillScheme) || element.Shape.GradientFill is not null)
                    throw new CodecException("invalid_presentation_fill", $"Presentation shape {element.Id} cannot combine image and solid or gradient fill state.");
                PptxImagePaintCodec.Validate(element.Shape.ImageFill, $"shape {element.Id} fill", assetCatalog);
            }
            PptxLineStyleCodec.Validate(element.Shape, element.Id);
            PptxShapeTransformCodec.Validate(element.Shape.Transform, element.Id);
            PptxShadowCodec.Validate(element.Shape.Shadow, element.Id);
            PptxGlowCodec.Validate(element.Shape.Glow, element.Id);
            PptxInnerShadowCodec.Validate(element.Shape.InnerShadow, element.Id);
            PptxReflectionCodec.Validate(element.Shape.Reflection, element.Id);
            PptxSoftEdgeCodec.Validate(element.Shape.SoftEdge, element.Id);
            PptxNonVisualAccessibilityCodec.Validate(element.Shape.Accessibility, element.Id);
            PptxTextCodec.Validate(element.Shape);
            foreach (var paragraph in element.Shape.TextBody?.Paragraphs ?? [])
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Image)
            PptxPictureCodec.Validate(element.Image, element.Id, assetCatalog, sourceBound: hasSourcePackage && element.Source is not null);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Table)
        {
            // Imported DrawingML tables may use a graphic-frame scale that
            // differs from the table grid coordinate space.  Keep authored
            // tables strict, while allowing a trusted source-bound table to
            // retain that bounded scale during a no-op or local edit.
            PptxTableCodec.Validate(
                element.Table,
                element.Id,
                allowScaledFrame: hasSourcePackage && element.Source is not null,
                assets: assetCatalog);
            items += checked((ulong)element.Table.Rows.Sum(row => row.Cells.Count));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Connector)
            PptxConnectorCodec.Validate(element.Connector, element.Id, element.Name);
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Chart)
        {
            // Imported source-bound charts may expose a bounded local
            // worksheet formula reference.  Source-free authoring remains
            // literal-only; the chart codec validates the formula profile
            // before allowing this narrow source-bound exception.
            PptxChartCodec.Validate(
                element.Chart,
                element.Id,
                element.Name,
                allowFormulas: hasSourcePackage && element.Source?.Editable == true);
            items += checked((ulong)(
                element.Chart.Series.Sum(series => series.Values.Count) +
                element.Chart.ComboSeries.Sum(entry => entry.Series?.Values.Count ?? 0)));
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Diagram)
        {
            PptxSmartArtCodec.Validate(element.Diagram, element.Id, assetCatalog);
            items += checked((ulong)(element.Diagram.Nodes.Count + element.Diagram.Connections.Count));
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Media)
        {
            if (hasSourcePackage)
                throw new CodecException("unsupported_presentation_authored_overlay", $"Presentation media {element.Id} cannot be added to a source-bound slide.");
            PptxMediaCodec.Validate(element.Media, element.Id, assetCatalog);
        }
        else if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
        {
            var group = element.Group;
            PptxNonVisualAccessibilityCodec.Validate(group.Accessibility, element.Id, "group");
            PptxFrameTransformCodec.Validate(group.FrameTransform, element.Id, "group");
            // DrawingML permits a group frame to start just outside the slide
            // canvas. Imported source-bound groups use that placement for
            // intentional bleed/trim and must remain editable without
            // normalizing the source coordinate. Keep source-free authoring
            // strict so a caller cannot accidentally create an invalid frame.
            var sourceBoundGroup = hasSourcePackage && element.Source is not null;
            if ((!sourceBoundGroup && (group.LeftEmu < 0 || group.TopEmu < 0)) ||
                group.WidthEmu <= 0 || group.HeightEmu <= 0 ||
                group.ChildWidthEmu <= 0 || group.ChildHeightEmu <= 0 || group.Children.Count == 0)
                throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} requires positive outer/child extents and at least one child.");
            var childIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var child in group.Children)
            {
                if (!childIds.Add(child.Id))
                    throw new CodecException("invalid_presentation_group", $"Presentation group {element.Id} contains duplicate child ID {child.Id}.");
                if (child.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                {
                    // An imported group may contain a vendor/native child that
                    // is not projected semantically.  It is safe to retain
                    // that child when the request carries its source binding;
                    // ApplyGroup separately proves whether a direct frame edit
                    // is allowed.  Source-free opaque children remain rejected
                    // so authoring cannot smuggle an unowned XML subtree into a
                    // group.
                    if (!hasSourcePackage || child.Source is null || string.IsNullOrWhiteSpace(child.Source.ElementSha256))
                        throw new CodecException("unsupported_presentation_features", $"Presentation group {element.Id} contains a source-free opaque child.");
                    if (child.Source.Editable && !HasValidSourceBoundNativeFrame(child.Opaque, allowNegativeOffset))
                        throw new CodecException("invalid_presentation_frame", $"Presentation native child {child.Id} has an invalid frame.");
                    continue;
                }
                ValidatePresentationElement(child, hasSourcePackage, assetCatalog, limits, ref items, depth + 1, allowNegativeOffset: true);
            }
        }
        else if (element.ContentCase != PresentationElement.ContentOneofCase.Opaque)
            throw new CodecException("missing_presentation_element_content", $"Presentation element {element.Id} has no content.");
        else if (element.Source?.Editable == true)
        {
            if (!HasValidSourceBoundNativeFrame(element.Opaque))
                throw new CodecException("invalid_presentation_frame", $"Presentation native object {element.Id} has an invalid frame.");
        }
    }

    private static bool HasValidSourceBoundNativeFrame(PresentationOpaqueElement frame, bool allowNegativeOffset = false)
    {
        var negativeOffsetAllowed = allowNegativeOffset || frame.NativeKind is "picture" or "connector" or "cxnSp";
        if ((!negativeOffsetAllowed && (frame.LeftEmu < 0 || frame.TopEmu < 0)) ||
            frame.WidthEmu < 0 || frame.HeightEmu < 0)
            return false;

        // DrawingML connectors may have a zero width or height when the two
        // endpoints share an axis.  The line geometry and connection targets
        // remain the source-bound payload; this operation only updates the
        // direct frame.  Require at least one positive extent so an invisible
        // zero-by-zero object is not promoted to an editable placement root.
        // Older imported envelopes may omit NativeKind on a nested opaque
        // connector; its source-bound capability has already been proven by
        // the native catalog before this validation runs.  Treat the empty
        // and native element-name forms as the same connector profile.
        if (frame.NativeKind is "connector" or "cxnSp" or "")
            return PptxConnectorCodec.IsFrame(frame.LeftEmu, frame.TopEmu, frame.WidthEmu, frame.HeightEmu) &&
                (frame.WidthEmu > 0 || frame.HeightEmu > 0);

        return frame.WidthEmu > 0 && frame.HeightEmu > 0;
    }

    private static void ValidatePlaceholders(
        string ownerId,
        IList<PresentationPlaceholder> placeholders,
        PptxAssetCatalog assetCatalog,
        EffectiveCodecLimits limits,
        ref ulong items)
    {
        if (placeholders.Count > 128)
            throw new CodecException("presentation_placeholder_budget_exceeded", $"Presentation owner {ownerId} exceeds the 128-placeholder budget.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var placeholder in placeholders)
        {
            items++;
            if (items > limits.MaxCells)
                throw new CodecException("presentation_item_budget_exceeded", $"Presentation exceeds max_cells semantic-item budget ({limits.MaxCells}).");
            if (!ids.Add(placeholder.Id))
                throw new CodecException("invalid_presentation_placeholder", $"Presentation owner {ownerId} contains duplicate placeholder ID {placeholder.Id}.");
            PptxPlaceholderCodec.Validate(placeholder);
            foreach (var paragraph in (placeholder.TextBody?.Paragraphs ?? []).Concat(placeholder.TextBody?.ListStyles ?? []))
                if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.PictureBullet &&
                    paragraph.PictureBullet.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
                    _ = assetCatalog.Get(paragraph.PictureBullet.AssetId);
        }
    }

    private static IReadOnlyList<string> AssertPackagePartsUnchangedExcept(
        byte[] sourceBytes,
        byte[] outputBytes,
        HashSet<string> allowedPaths)
    {
        var before = PackagePartHashes(sourceBytes);
        var after = PackagePartHashes(outputBytes);
        var inventoryChanges = before.Keys.Concat(after.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !before.ContainsKey(path) || !after.ContainsKey(path))
            .Where(path => !allowedPaths.Contains(path))
            .Take(8)
            .ToArray();
        if (inventoryChanges.Length > 0)
            throw new CodecException("presentation_package_topology_changed", $"Source-preserving PPTX export changed unowned OPC part inventory: {string.Join(", ", inventoryChanges)}.");
        var changed = before.Keys.Intersect(after.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(path => !before[path].Equals(after[path], StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var unexpected = changed.Where(path => !allowedPaths.Contains(path)).Take(8).ToArray();
        if (unexpected.Length > 0)
            throw new CodecException("presentation_unowned_part_changed", $"Source-preserving PPTX export changed unowned package parts: {string.Join(", ", unexpected)}.");
        // Include owned package additions in the receipt.  An image/SVG
        // replacement changes its slide relationship and creates a new media
        // part; both are part of the observable source-bound package delta.
        var added = after.Keys.Except(before.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(path => allowedPaths.Contains(path));
        return changed.Concat(added).Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static void AssertOpaqueGraphMatchesWithModeledAdditions(
        OpaqueOpcGraph expected,
        OpaqueOpcGraph actual,
        IReadOnlySet<string> allowedAddedRelationshipIds,
        IReadOnlySet<string> allowedAddedPartPaths,
        IReadOnlyDictionary<string, string> allowedChangedPartHashes,
        IReadOnlySet<string> removedSourcePartPaths,
        IReadOnlySet<string> removedSourceRelationshipKeys)
    {
        var guarded = actual.Clone();
        var modeledRelationshipPartPaths = allowedAddedRelationshipIds
            .Concat(removedSourceRelationshipKeys)
            .Select(key => key.IndexOf('\0') is var separator && separator > 0
                ? RelationshipPartPathFor(key[..separator])
                : string.Empty)
            .Where(path => path.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var removed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in guarded.PackageRelationships.ToArray())
        {
            var key = $"{relationship.SourcePath}\0{relationship.Id}";
            if (!allowedAddedRelationshipIds.Contains(key)) continue;
            guarded.PackageRelationships.Remove(relationship);
            removed.Add(key);
        }
        if (!removed.SetEquals(allowedAddedRelationshipIds))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX relationship additions do not match the relationships written to the package.");
        var removedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in guarded.Parts.ToArray())
        {
            if (!allowedAddedPartPaths.Contains(part.Path)) continue;
            guarded.Parts.Remove(part);
            removedParts.Add(part.Path);
        }
        if (!removedParts.SetEquals(allowedAddedPartPaths))
            throw new CodecException("opaque_content_not_preserved", "Modeled PPTX additions do not match the parts written to the package.");
        foreach (var (path, requestedHash) in allowedChangedPartHashes)
        {
            var before = expected.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            var after = guarded.Parts.SingleOrDefault(part => part.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            var expectedRelationships = before?.Relationships
                .Where(relationship => !removedSourceRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"))
                .ToArray();
            var actualRelationships = after?.Relationships
                .Where(relationship => !allowedAddedRelationshipIds.Contains($"{relationship.SourcePath}\0{relationship.Id}"))
                .ToArray();
            if (before is null || after is null ||
                !before.ContentType.Equals(after.ContentType, StringComparison.OrdinalIgnoreCase) ||
                !after.Sha256.Equals(requestedHash, StringComparison.OrdinalIgnoreCase) ||
                !(expectedRelationships ?? []).Select(OpaqueRelationshipSignature)
                    .SequenceEqual((actualRelationships ?? []).Select(OpaqueRelationshipSignature), StringComparer.Ordinal))
                throw new CodecException("opaque_content_not_preserved", $"Modeled PPTX OLE workbook replacement did not preserve the package contract for {path}.", path);
        }
        PackageGuards.AssertOpaqueGraphMatches(
            expected,
            guarded,
            "opaque_content_not_preserved",
            ignoreRelationship: relationship =>
                removedSourcePartPaths.Contains(relationship.SourcePath) ||
                removedSourceRelationshipKeys.Contains($"{relationship.SourcePath}\0{relationship.Id}"),
            ignorePart: part => allowedChangedPartHashes.ContainsKey(part.Path) ||
                modeledRelationshipPartPaths.Contains(part.Path) ||
                removedSourcePartPaths.Contains(part.Path));
    }

    private static string OpaqueRelationshipSignature(OpaqueOpcRelationship relationship) =>
        $"{relationship.SourcePath}\0{relationship.Id}\0{relationship.Type}\0{relationship.Target}\0{relationship.TargetMode}";

    private static string RelationshipPartPathFor(string ownerPath)
    {
        var separator = ownerPath.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : ownerPath[..separator];
        var fileName = separator < 0 ? ownerPath : ownerPath[(separator + 1)..];
        return directory.Length == 0 ? $"_rels/{fileName}.rels" : $"{directory}/_rels/{fileName}.rels";
    }

    private static void AssertPlannedPartsRemoved(byte[] sourceBytes, byte[] outputBytes, IReadOnlySet<string> removedPartPaths)
    {
        if (removedPartPaths.Count == 0) return;
        var before = PackagePartHashes(sourceBytes);
        var after = PackagePartHashes(outputBytes);
        var retained = removedPartPaths
            .Where(path => before.ContainsKey(path) && after.ContainsKey(path))
            .Take(8)
            .ToArray();
        if (retained.Length > 0)
            throw new CodecException(
                "presentation_delete_incomplete",
                $"Source-preserving PPTX deletion retained planned package parts: {string.Join(", ", retained)}.");
    }

    private static bool IsNumberedSlidePath(string path)
    {
        const string prefix = "ppt/slides/slide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedMasterPath(string path)
    {
        const string prefix = "ppt/slideMasters/slideMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedLayoutPath(string path)
    {
        const string prefix = "ppt/slideLayouts/slideLayout";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlidePath(string path)
    {
        const string prefix = "ppt/notesSlides/notesSlide";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesSlideRelationshipPath(string path)
    {
        const string prefix = "ppt/notesSlides/_rels/notesSlide";
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesMasterPath(string path)
    {
        const string prefix = "ppt/notesMasters/notesMaster";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedNotesMasterRelationshipPath(string path)
    {
        const string prefix = "ppt/notesMasters/_rels/notesMaster";
        const string suffix = ".xml.rels";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsNumberedThemePath(string path)
    {
        const string suffix = ".xml";
        foreach (var prefix in new[] { "ppt/theme/theme", "ppt/slideMasters/theme/theme" })
            if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
                path[prefix.Length..^suffix.Length].Length > 0 &&
                path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit))
                return true;
        return false;
    }

    private static bool IsNumberedCommentsPath(string path)
    {
        const string prefix = "ppt/comments/comment";
        const string suffix = ".xml";
        return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
               path[prefix.Length..^suffix.Length].Length > 0 &&
               path[prefix.Length..^suffix.Length].All(char.IsAsciiDigit);
    }

    private static bool IsLegacyCommentAuthorsPath(string path) =>
        path.Equals("ppt/commentAuthors.xml", StringComparison.OrdinalIgnoreCase);

    private static void ValidatePreservedSlideElements(
        byte[] sourceBytes,
        byte[] outputBytes,
        PresentationArtifact requested,
        EffectiveCodecLimits limits,
        IReadOnlySet<string> changedPartPaths)
    {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var orderedSourceSlides = OrderedSlideParts(sourcePackage);
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (sourceTargets.Length != requested.Slides.Count || outputSlides.Length != requested.Slides.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide topology changed during source-preserving export.");
        var retainedSourceSlideParts = sourceTargets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            var outputSlide = outputSlides[targetIndex];
            if (!target.IsClone && !PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
            if (target.IsClone)
            {
                if (PartPath(outputSlide).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_postwrite_clone_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source slide copy.", PartPath(outputSlide));
                if (target.Target.ElementDeletions.Count == 0)
                {
                    if (!HashElement(target.Source.Part.Slide!).Equals(HashElement(outputSlide.Slide!), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_clone_mismatch", $"PPTX clone {targetIndex + 1} is not an independent exact source slide copy.", PartPath(outputSlide));
                    PptxSlideCloneCodec.Validate(target.Source, outputSlide, retainedSourceSlideParts);
                }
                else
                {
                    ValidateCloneElementProjection(target.Source, outputSlide, target.Target);
                }
            }
            else
            {
                PptxSpeakerNotesCodec.ValidateSourceBoundOutput(
                    sourcePresentationPart,
                    outputPresentationPart,
                    target.Source.Part,
                    outputSlide,
                    target.Target,
                    targetIndex);
                PptxLegacyCommentsCodec.ValidateSourceBoundOutput(
                    sourcePresentationPart,
                    outputPresentationPart,
                    target.Source.Part,
                    outputSlide,
                    target.Target,
                    targetIndex);
            }
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputIdByPartPath = retainedTargets
            .Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits, retainImportedAssetData: false);
        var outputAssets = new PptxAssetCatalog([], limits, retainImportedAssetData: false);
        var customShowCatalog = PptxCustomShowCatalog.From(requested.CustomShows);

        for (var slideIndex = 0; slideIndex < requested.Slides.Count; slideIndex++)
        {
            if (sourceTargets[slideIndex].IsClone) continue;
            var target = sourceTargets[slideIndex];
            var sourceSlide = target.Source.Part;
            var outputSlide = outputSlides[slideIndex];
            var outputRoot = outputSlide.Slide ??
                throw new CodecException("missing_slide_root", $"PPTX output slide {slideIndex + 1} has no slide root.", PartPath(outputSlide));
            var sourceRoot = sourceSlide.Slide!;
            var outputName = outputRoot.CommonSlideData?.Name?.Value ?? string.Empty;
            if (!string.Equals(outputName, requested.Slides[slideIndex].Name, StringComparison.Ordinal))
                throw new CodecException(
                    "presentation_postwrite_slide_name_mismatch",
                    $"PPTX slide {slideIndex + 1} name does not match the requested source-bound value.",
                    PartPath(outputSlide));
            if (!PptxSlideVisibilityCodec.Matches(requested.Slides[slideIndex], sourceRoot, outputRoot))
                throw new CodecException(
                    "presentation_postwrite_slide_visibility_mismatch",
                    $"PPTX slide {slideIndex + 1} visibility does not match the requested source-bound value.",
                    PartPath(outputSlide));
            var sourceTransition = PptxTransitionCodec.Read(sourceRoot);
            var outputTransition = PptxTransitionCodec.Read(outputRoot);
            var requestedTransition = requested.Slides[slideIndex].Transition;
            var transitionChanged = !PptxTransitionCodec.SemanticHash(requestedTransition)
                .Equals(PptxTransitionCodec.SemanticHash(sourceTransition), StringComparison.OrdinalIgnoreCase);
            if (!transitionChanged && !PptxTransitionCodec.ElementHash(sourceRoot)
                    .Equals(PptxTransitionCodec.ElementHash(outputRoot), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unchanged_transition_modified",
                    $"PPTX slide {slideIndex + 1} unchanged transition was modified during export.",
                    PartPath(outputSlide));
            if (!PptxTransitionCodec.SemanticHash(outputTransition)
                    .Equals(PptxTransitionCodec.SemanticHash(requestedTransition), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_transition_semantics_mismatch",
                    $"PPTX slide {slideIndex + 1} transition does not match requested semantics after export.",
                    PartPath(outputSlide));

            // AssertPackagePartsUnchangedExcept has already proved every
            // untouched OPC part byte-for-byte identical. When the exporter
            // changed only numbered slide XML parts, keep the cheap topology,
            // name, visibility, and transition checks above, then skip the
            // element/animation/background DOM walk for untouched slides.
            // Relationship, media, master/layout, and global changes stay on
            // the conservative full validation path.
            var slideOnlyChanges = changedPartPaths.Count > 0 &&
                changedPartPaths.All(IsNumberedSlidePath);
            if (slideOnlyChanges && !changedPartPaths.Contains(PartPath(outputSlide)))
                continue;

            var sourceTimingElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
                throw new CodecException("missing_shape_tree", $"PPTX source slide {slideIndex + 1} has no shape tree.", PartPath(sourceSlide)));
            var sourceTimingIds = NativeElementIds(sourceTimingElements, requested.Slides[slideIndex].Id);
            P.Slide? previousSourceRoot = null;
            IReadOnlyDictionary<uint, string>? previousSourceTimingIds = null;
            string? previousSourceSlideId = null;
            if (target.Source.Index > 0 && PptxTimingCodec.HasMorph(sourceRoot))
            {
                previousSourceRoot = orderedSourceSlides[target.Source.Index - 1].Slide;
                var previousSourceTarget = sourceTargets.FirstOrDefault(candidate =>
                    !candidate.IsClone && candidate.Source.Index == target.Source.Index - 1);
                previousSourceSlideId = previousSourceTarget?.Target.Id;
                var previousSourceTree = previousSourceRoot?.CommonSlideData?.ShapeTree;
                if (previousSourceTree is not null && !string.IsNullOrWhiteSpace(previousSourceSlideId))
                    previousSourceTimingIds = NativeElementIds(ShapeElements(previousSourceTree), previousSourceSlideId);
            }
            var sourceTiming = PptxTimingCodec.Read(
                sourceRoot,
                sourceTimingIds,
                previousSourceRoot,
                previousSourceTimingIds,
                previousSourceSlideId);
            P.Slide? previousOutputRoot = null;
            IReadOnlyDictionary<uint, string>? previousOutputTimingIds = null;
            string? previousOutputSlideId = null;
            if (slideIndex > 0 && PptxTimingCodec.HasMorph(outputRoot))
            {
                previousOutputRoot = outputSlides[slideIndex - 1].Slide;
                previousOutputSlideId = requested.Slides[slideIndex - 1].Id;
                var previousOutputTree = previousOutputRoot?.CommonSlideData?.ShapeTree;
                if (previousOutputTree is not null)
                    previousOutputTimingIds = NativeElementIds(ShapeElements(previousOutputTree), previousOutputSlideId);
            }
            var outputTiming = PptxTimingCodec.Read(
                outputRoot,
                sourceTimingIds,
                previousOutputRoot,
                previousOutputTimingIds,
                previousOutputSlideId);
            var requestedTimingHash = PptxTimingCodec.SemanticHash(requested.Slides[slideIndex].Animations, requested.Slides[slideIndex].Morph);
            var requestedOpaqueNoop = sourceTiming.Present && !sourceTiming.Editable && requested.Slides[slideIndex].Animations.Count == 0 && requested.Slides[slideIndex].Morph is null;
            if (!requestedOpaqueNoop && !outputTiming.SemanticSha256.Equals(requestedTimingHash, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_timing_semantics_mismatch",
                    $"PPTX slide {slideIndex + 1} timing does not match requested animation semantics after export.",
                    PartPath(outputSlide));
            var sourceContext = new PptxPartContext(sourceSlide, sourceIdByPartPath, assets: sourceAssets, customShows: customShowCatalog, slideNumber: checked(slideIndex + 1));
            var outputContext = new PptxPartContext(outputSlides[slideIndex], outputIdByPartPath, assets: outputAssets, customShows: customShowCatalog, slideNumber: checked(slideIndex + 1));
            var outputBackground = PptxBackgroundCodec.Read(outputRoot.CommonSlideData, outputContext);
            if (!BackgroundSemanticHash(outputBackground).Equals(BackgroundSemanticHash(requested.Slides[slideIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_slide_background_mismatch",
                    $"PPTX slide {slideIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputSlide));
            var before = ShapeElements(sourceSlide.Slide!.CommonSlideData!.ShapeTree!);
            var after = ShapeElements(outputRoot.CommonSlideData!.ShapeTree!);
            var (elements, authoredElements) = SplitSourceBoundElements(requested.Slides[slideIndex], before.Length, slideIndex, outputSlide);
            var authoredNativeIds = AuthoredOverlayNativeIds(before, authoredElements, slideIndex, outputSlide);
            var outputNativeIdSet = after.SelectMany(PptxElementDeletionCodec.NativeIds).ToHashSet();
            var afterIds = sourceTimingIds
                .Where(entry => outputNativeIdSet.Contains(entry.Key))
                .ToDictionary(entry => entry.Key, entry => entry.Value);
            foreach (var (elementId, nativeId) in authoredNativeIds)
                afterIds.Add(nativeId, elementId);
            var deletions = requested.Slides[slideIndex].ElementDeletions;
            if (before.Length != elements.Length + deletions.Count || after.Length != elements.Length + authoredElements.Length)
                throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} element topology changed during source-preserving export.", PartPath(outputSlide));
            var outputNativeIds = outputNativeIdSet;
            foreach (var deletion in deletions)
            {
                var binding = deletion.Source ?? throw new CodecException(
                    "missing_presentation_element_deletion_binding",
                    $"PPTX slide {slideIndex + 1} element deletion {deletion.Id} lost its source binding after export.",
                    PartPath(outputSlide));
                var sourceElementIndex = checked((int)binding.ShapeTreeIndex);
                if (sourceElementIndex >= before.Length ||
                    !binding.ElementSha256.Equals(HashElement(before[sourceElementIndex]), StringComparison.OrdinalIgnoreCase) ||
                    PptxElementDeletionCodec.NativeIds(before[sourceElementIndex]).Overlaps(outputNativeIds))
                    throw new CodecException(
                        "presentation_postwrite_element_delete_mismatch",
                        $"PPTX slide {slideIndex + 1} deleted element {sourceElementIndex + 1} remains or no longer matches its source binding.",
                        PartPath(outputSlide));
                var plan = PptxElementDeletionCodec.Analyze(sourceSlide, before[sourceElementIndex], before);
                if (!plan.Supported)
                    throw new CodecException(
                        "presentation_postwrite_element_delete_mismatch",
                        $"PPTX slide {slideIndex + 1} deleted element {sourceElementIndex + 1} no longer satisfies its source deletion proof.",
                        PartPath(outputSlide));
            }
            for (var elementIndex = 0; elementIndex < elements.Length; elementIndex++)
            {
                var request = elements[elementIndex];
                outputContext.DeriveAutomaticFields = ContainsAutomaticFields(request);
                var binding = request.Source!;
                var sourceElementIndex = checked((int)binding.ShapeTreeIndex);
                if (sourceElementIndex >= before.Length)
                    throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} retained element {elementIndex + 1} has an invalid source index.", PartPath(outputSlide));
                var beforeElement = before[sourceElementIndex];
                var afterElement = after[elementIndex];
                var changed = HasParagraphTabStopRemovalIntent(request) || HasParagraphRightMarginRemovalIntent(request) || !SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
                if (!changed)
                {
                    if (!HashElement(beforeElement).Equals(HashElement(afterElement), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unchanged_element_modified",
                            $"PPTX slide {slideIndex + 1} unchanged retained element {sourceElementIndex + 1} was modified during export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Opaque)
                {
                    if (!NativeObjectResidualHash(beforeElement).Equals(NativeObjectResidualHash(afterElement), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_native_content_changed",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputFrame = ReadFrame(afterElement);
                    if (!ElementName(afterElement, elementIndex).Equals(request.Name, StringComparison.Ordinal) ||
                        outputFrame.Left != request.Opaque.LeftEmu || outputFrame.Top != request.Opaque.TopEmu ||
                        outputFrame.Width != request.Opaque.WidthEmu || outputFrame.Height != request.Opaque.HeightEmu)
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited native object {elementIndex + 1} does not match the requested name/frame.",
                            PartPath(outputSlides[slideIndex]));
                    PptxDiagramTextCodec.ValidateSourceBoundOutput(
                        sourceSlide,
                        outputSlide,
                        beforeElement,
                        afterElement,
                        request.Opaque);
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Diagram)
                {
                    if (beforeElement is not P.GraphicFrame || afterElement is not P.GraphicFrame afterDiagramFrame)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited SmartArt {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    var outputDiagram = ReadElement(afterDiagramFrame, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (outputDiagram.ContentCase != PresentationElement.ContentOneofCase.Diagram ||
                        !SemanticHash(outputDiagram).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited SmartArt {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Group)
                {
                    if (afterElement is not P.GroupShape afterGroup)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (beforeElement is P.GraphicFrame)
                    {
                        var sourceDiagram = ReadElement(
                            beforeElement,
                            slideIndex,
                            sourceElementIndex,
                            sourceContext,
                            elementIdsByNativeId: sourceTimingIds);
                        if (sourceDiagram.ContentCase != PresentationElement.ContentOneofCase.Diagram ||
                            !sourceDiagram.Diagram.DrawingCacheVerified ||
                            !sourceDiagram.Diagram.Drawing.Equals(request.Group))
                            throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} group {elementIndex + 1} is not an exact verified SmartArt cache detachment.", PartPath(outputSlides[slideIndex]));
                        var detachedSemantic = ReadElement(afterGroup, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                        if (!SemanticHash(detachedSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                            throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} detached SmartArt group {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                        continue;
                    }
                    if (beforeElement is not P.GroupShape beforeGroup)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    ValidateGroupOutput(beforeGroup, afterGroup, request, sourceContext, outputContext, afterIds, slideIndex, $"element {elementIndex + 1}");
                    var outputGroupSemantic = ReadElement(afterGroup, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputGroupSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited group {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Image)
                {
                    if (beforeElement is not P.Picture beforePicture || afterElement is not P.Picture afterPicture)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_picture_content_changed",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputPictureSemantic = ReadElement(afterPicture, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputPictureSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited image {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Table)
                {
                    if (beforeElement is not P.GraphicFrame beforeTable || afterElement is not P.GraphicFrame afterTable)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!TableResidualHash(beforeTable, sourceContext).Equals(TableResidualHash(afterTable, outputContext), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_unmodeled_table_content_changed",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} changed unmodeled native content.",
                            PartPath(outputSlides[slideIndex]));
                    var outputTableSemantic = ReadElement(afterTable, slideIndex, elementIndex, outputContext);
                    if (!SemanticHash(outputTableSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "presentation_postwrite_semantics_mismatch",
                            $"PPTX slide {slideIndex + 1} edited table {elementIndex + 1} does not match requested semantics after export.",
                            PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Connector)
                {
                    if (beforeElement is not P.ConnectionShape beforeConnector || afterElement is not P.ConnectionShape afterConnector)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} changed unmodeled native content.", PartPath(outputSlides[slideIndex]));
                    var outputConnectorSemantic = ReadElement(afterConnector, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputConnectorSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited connector {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (request.ContentCase == PresentationElement.ContentOneofCase.Chart)
                {
                    if (beforeElement is not P.GraphicFrame beforeChart || afterElement is not P.GraphicFrame afterChart)
                        throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                    if (!ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} changed unmodeled frame content.", PartPath(outputSlides[slideIndex]));
                    var outputChartSemantic = ReadElement(afterChart, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                    if (!SemanticHash(outputChartSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} edited chart {elementIndex + 1} does not match requested semantics after export.", PartPath(outputSlides[slideIndex]));
                    continue;
                }
                if (beforeElement is not P.Shape beforeShape || afterElement is not P.Shape afterShape)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} edited element {elementIndex + 1} changed native element type.", PartPath(outputSlides[slideIndex]));
                if (!ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_unmodeled_shape_content_changed",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} changed unmodeled native content.",
                        PartPath(outputSlides[slideIndex]));
                var outputSemantic = ReadElement(afterShape, slideIndex, elementIndex, outputContext, elementIdsByNativeId: afterIds);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} edited shape {elementIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlides[slideIndex]));
            }
            for (var authoredIndex = 0; authoredIndex < authoredElements.Length; authoredIndex++)
            {
                var request = authoredElements[authoredIndex];
                outputContext.DeriveAutomaticFields = ContainsAutomaticFields(request);
                var outputIndex = elements.Length + authoredIndex;
                var outputElement = after[outputIndex];
                var authoredOutputNativeIds = PptxElementDeletionCodec.NativeIds(outputElement).ToArray();
                var nativeTypeMatches = request.ContentCase switch
                {
                    PresentationElement.ContentOneofCase.Shape => outputElement is P.Shape,
                    PresentationElement.ContentOneofCase.Image => outputElement is P.Picture,
                    _ => false,
                };
                if (!nativeTypeMatches || authoredOutputNativeIds.Length != 1 || authoredOutputNativeIds[0] != authoredNativeIds[request.Id])
                    throw new CodecException(
                        "presentation_postwrite_authored_overlay_mismatch",
                        $"PPTX slide {slideIndex + 1} authored overlay {authoredIndex + 1} changed native type or identity during export.",
                        PartPath(outputSlide));
                var outputSemantic = ReadElement(outputElement, slideIndex, outputIndex, outputContext, elementIdsByNativeId: afterIds);
                if (!SemanticHash(outputSemantic).Equals(SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "presentation_postwrite_semantics_mismatch",
                        $"PPTX slide {slideIndex + 1} authored overlay {authoredIndex + 1} does not match requested semantics after export.",
                        PartPath(outputSlide));
            }
        }
    }

    private static void ValidateCloneElementProjection(
        PptxSourceSlideEntry source,
        SlidePart output,
        PresentationSlide requested)
    {
        var sourceRoot = source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
        var outputRoot = output.Slide ??
            throw new CodecException("missing_slide_root", "Presentation cloned slide has no slide root.", PartPath(output));
        var sourceElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {source.Index + 1} has no shape tree.", PartPath(source.Part)));
        var outputElements = ShapeElements(outputRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", "Presentation cloned slide has no shape tree.", PartPath(output)));
        if (sourceElements.Length != requested.Elements.Count + requested.ElementDeletions.Count || outputElements.Length != requested.Elements.Count)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX component clone output does not match its retained/deleted element projection.", PartPath(output));
        var outputNativeIdList = outputElements.SelectMany(PptxElementDeletionCodec.NativeIds).ToArray();
        if (outputNativeIdList.GroupBy(id => id).Any(group => group.Count() > 1))
            throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone retained duplicate native drawing IDs.", PartPath(output));
        var outputNativeIds = outputNativeIdList.ToHashSet();
        foreach (var deletion in requested.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", "PPTX component clone deletion lost its source binding.", PartPath(output));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length ||
                !binding.ElementSha256.Equals(HashElement(sourceElements[sourceIndex]), StringComparison.OrdinalIgnoreCase) ||
                PptxElementDeletionCodec.NativeIds(sourceElements[sourceIndex]).Overlaps(outputNativeIds))
                throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone retained a deleted source element or lost its source proof.", PartPath(output));
            var plan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[sourceIndex], sourceElements, allowDuplicateNativeIds: true);
            if (!plan.Supported)
                throw new CodecException("presentation_postwrite_element_delete_mismatch", "PPTX component clone deletion no longer satisfies its source deletion proof.", PartPath(output));
        }
        for (var retainedIndex = 0; retainedIndex < requested.Elements.Count; retainedIndex++)
        {
            var binding = requested.Elements[retainedIndex].Source ??
                throw new CodecException("missing_presentation_element_binding", "PPTX component clone retained element lost its source binding.", PartPath(output));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length ||
                !binding.ElementSha256.Equals(HashElement(sourceElements[sourceIndex]), StringComparison.OrdinalIgnoreCase) ||
                !HashElement(sourceElements[sourceIndex]).Equals(HashElement(outputElements[retainedIndex]), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_clone_mismatch", "PPTX component clone changed an element that was not selected for reuse.", PartPath(output));
        }
    }

    private static void ValidateGroupOutput(
        P.GroupShape before,
        P.GroupShape after,
        PresentationElement request,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        IReadOnlyDictionary<uint, string> afterIds,
        int slideIndex,
        string location)
    {
        if (!GroupShellResidualHash(before).Equals(GroupShellResidualHash(after), StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_unmodeled_group_content_changed", $"PPTX slide {slideIndex + 1} {location} changed unmodeled group-shell content.", PartPath(outputContext.Owner));
        var beforeChildren = GroupElements(before);
        var afterChildren = GroupElements(after);
        if (beforeChildren.Length != request.Group.Children.Count || afterChildren.Length != request.Group.Children.Count)
            throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} {location} group topology changed during export.", PartPath(outputContext.Owner));
        var requestedSourceOrder = new int[request.Group.Children.Count];
        var requestedSourceIndexes = new HashSet<int>();
        for (var requestedIndex = 0; requestedIndex < request.Group.Children.Count; requestedIndex++)
        {
            var binding = request.Group.Children[requestedIndex].Source ??
                throw new CodecException("missing_presentation_element_binding", $"PPTX slide {slideIndex + 1} {location} child {requestedIndex + 1} is missing its source binding.", PartPath(outputContext.Owner));
            if (binding.ShapeTreeIndex >= (uint)beforeChildren.Length || !requestedSourceIndexes.Add((int)binding.ShapeTreeIndex))
                throw new CodecException("presentation_postwrite_topology_changed", $"PPTX slide {slideIndex + 1} {location} child order is not a complete source permutation.", PartPath(outputContext.Owner));
            requestedSourceOrder[requestedIndex] = (int)binding.ShapeTreeIndex;
        }

        for (var index = 0; index < request.Group.Children.Count; index++)
        {
            var child = request.Group.Children[index];
            var binding = child.Source ?? throw new CodecException("missing_presentation_element_binding", $"PPTX slide {slideIndex + 1} {location} child {index + 1} is missing its source binding.", PartPath(outputContext.Owner));
            var sourceIndex = requestedSourceOrder[index];
            var beforeChild = beforeChildren[sourceIndex];
            var afterChild = afterChildren[index];
            var changed = HasParagraphTabStopRemovalIntent(child) || HasParagraphRightMarginRemovalIntent(child) || !SemanticHash(child).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                if (!HashElement(beforeChild).Equals(HashElement(afterChild), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_element_modified", $"PPTX slide {slideIndex + 1} {location} unchanged child {index + 1} was modified during export.", PartPath(outputContext.Owner));
                continue;
            }

            if (child.ContentCase == PresentationElement.ContentOneofCase.Group)
            {
                if (beforeChild is not P.GroupShape beforeGroup || afterChild is not P.GroupShape afterGroup)
                    throw new CodecException("presentation_postwrite_element_mismatch", $"PPTX slide {slideIndex + 1} {location} child group {index + 1} changed native element type.", PartPath(outputContext.Owner));
                ValidateGroupOutput(beforeGroup, afterGroup, child, sourceContext, outputContext, afterIds, slideIndex, $"{location} child {index + 1}");
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Image)
            {
                if (beforeChild is not P.Picture beforePicture || afterChild is not P.Picture afterPicture ||
                    !PictureResidualHash(beforePicture).Equals(PictureResidualHash(afterPicture), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_picture_content_changed", $"PPTX slide {slideIndex + 1} {location} child image {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Table)
            {
                if (beforeChild is not P.GraphicFrame beforeTable || afterChild is not P.GraphicFrame afterTable ||
                    !TableResidualHash(beforeTable, sourceContext).Equals(TableResidualHash(afterTable, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_table_content_changed", $"PPTX slide {slideIndex + 1} {location} child table {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Connector)
            {
                if (beforeChild is not P.ConnectionShape beforeConnector || afterChild is not P.ConnectionShape afterConnector ||
                    !ConnectorResidualHash(beforeConnector).Equals(ConnectorResidualHash(afterConnector), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_connector_content_changed", $"PPTX slide {slideIndex + 1} {location} child connector {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Chart)
            {
                if (beforeChild is not P.GraphicFrame beforeChart || afterChild is not P.GraphicFrame afterChart ||
                    !ChartFrameResidualHash(beforeChart).Equals(ChartFrameResidualHash(afterChart), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_chart_frame_changed", $"PPTX slide {slideIndex + 1} {location} child chart {index + 1} changed unmodeled frame content.", PartPath(outputContext.Owner));
            }
            else if (child.ContentCase == PresentationElement.ContentOneofCase.Opaque)
            {
                // Opaque children can have a separately proven direct-frame
                // edit (for example a vendor group or content part).  Keep
                // their descendants and relationships source-bound while
                // checking only the frame/name that this bounded operation is
                // allowed to change.
                if (!NativeObjectResidualHash(beforeChild).Equals(NativeObjectResidualHash(afterChild), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_native_content_changed", $"PPTX slide {slideIndex + 1} {location} child native object {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
                var outputFrame = ReadFrame(afterChild);
                if (!ElementName(afterChild, index).Equals(child.Name, StringComparison.Ordinal) ||
                    outputFrame.Left != child.Opaque.LeftEmu || outputFrame.Top != child.Opaque.TopEmu ||
                    outputFrame.Width != child.Opaque.WidthEmu || outputFrame.Height != child.Opaque.HeightEmu)
                    throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} {location} child native object {index + 1} does not match requested name/frame.", PartPath(outputContext.Owner));
            }
            else
            {
                if (beforeChild is not P.Shape beforeShape || afterChild is not P.Shape afterShape ||
                    !ShapeResidualHash(beforeShape, sourceContext).Equals(ShapeResidualHash(afterShape, outputContext), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unmodeled_shape_content_changed", $"PPTX slide {slideIndex + 1} {location} child shape {index + 1} changed unmodeled content.", PartPath(outputContext.Owner));
            }

            var outputSemantic = ReadElement(afterChild, request.Id, index, outputContext, elementIdsByNativeId: afterIds);
            if (!SemanticHash(outputSemantic).Equals(SemanticHash(child), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_semantics_mismatch", $"PPTX slide {slideIndex + 1} {location} child {index + 1} does not match requested semantics after export.", PartPath(outputContext.Owner));
        }
    }

    private static string GroupShellResidualHash(P.GroupShape source)
    {
        var clone = (P.GroupShape)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(clone);
        foreach (var child in GroupElements(clone)) child.Remove();
        if (clone.NonVisualGroupShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (clone.GroupShapeProperties?.GetFirstChild<A.TransformGroup>() is { } transform)
        {
            transform.Offset!.X = 0L;
            transform.Offset.Y = 0L;
            transform.Extents!.Cx = 1L;
            transform.Extents.Cy = 1L;
            transform.ChildOffset!.X = 0L;
            transform.ChildOffset.Y = 0L;
            transform.ChildExtents!.Cx = 1L;
            transform.ChildExtents.Cy = 1L;
            PptxFrameTransformCodec.Scrub(transform);
        }
        return HashElement(clone);
    }

    private static void ValidatePreservedMasterAndLayoutContent(
        byte[] sourceBytes,
        byte[] outputBytes,
        PresentationArtifact requested,
        EffectiveCodecLimits limits,
        IReadOnlySet<string> changedPartPaths)
    {
        // A source-bound edit that touched only slide XML cannot have changed
        // any master/layout bytes: the package-wide part hash assertion above
        // has already proved that. Avoid opening a second pair of large
        // PresentationDocument DOMs for this common build path. Keep the full
        // validator for no-op, global, relationship, media, master, and layout
        // changes, where those bytes can participate in the semantic contract.
        if (changedPartPaths.Count > 0 && changedPartPaths.All(IsNumberedSlidePath))
            return;

        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var outputStream = new MemoryStream(outputBytes, writable: false);
        using var sourcePackage = PresentationDocument.Open(sourceStream, isEditable: false);
        using var outputPackage = PresentationDocument.Open(outputStream, isEditable: false);
        var sourcePresentationPart = sourcePackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX source package has no Presentation part.", "ppt/presentation.xml");
        var outputPresentationPart = outputPackage.PresentationPart ??
            throw new CodecException("missing_presentation_part", "PPTX output package has no Presentation part.", "ppt/presentation.xml");
        var sourceGraph = ReadMasterGraph(sourcePresentationPart);
        var outputGraph = ReadMasterGraph(outputPresentationPart);
        if (sourceGraph.Length != requested.Masters.Count || outputGraph.Length != requested.Masters.Count)
            throw new CodecException("presentation_postwrite_master_topology_changed", "PPTX master topology changed during source-preserving export.");
        var sourceTargets = BindSourcePreservingSlides(
            sourcePresentationPart,
            sourcePresentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [],
            requested.Slides);
        var outputSlides = OrderedSlideParts(outputPackage);
        if (outputSlides.Length != sourceTargets.Length)
            throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        for (var targetIndex = 0; targetIndex < sourceTargets.Length; targetIndex++)
        {
            var target = sourceTargets[targetIndex];
            if (!target.IsClone && !PartPath(outputSlides[targetIndex]).Equals(PartPath(target.Source.Part), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_topology_changed", "PPTX slide order does not match the requested source-bound order.", "ppt/presentation.xml");
        }
        var retainedTargets = sourceTargets.Where(target => !target.IsClone).ToArray();
        var sourceSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var outputSlideMap = retainedTargets.Select(target => (Path: PartPath(target.Source.Part), Id: target.Target.Id))
            .ToDictionary(item => item.Path, item => item.Id, StringComparer.OrdinalIgnoreCase);
        var sourceAssets = new PptxAssetCatalog([], limits, retainImportedAssetData: false);
        var outputAssets = new PptxAssetCatalog([], limits, retainImportedAssetData: false);
        var customShowCatalog = PptxCustomShowCatalog.From(requested.CustomShows);
        for (var masterIndex = 0; masterIndex < requested.Masters.Count; masterIndex++)
        {
            var before = sourceGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX source master {masterIndex + 1} has no root.");
            var after = outputGraph[masterIndex].Part.SlideMaster ??
                throw new CodecException("missing_slide_master_root", $"PPTX output master {masterIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceGraph[masterIndex].Part, sourceSlideMap, assets: sourceAssets, customShows: customShowCatalog);
            var outputContext = new PptxPartContext(outputGraph[masterIndex].Part, outputSlideMap, assets: outputAssets, customShows: customShowCatalog);
            if (!MasterResidualHash(before, sourceContext).Equals(MasterResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_master_content_changed",
                    $"PPTX master {masterIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputGraph[masterIndex].Part));
            var outputStyles = PptxMasterTextStylesCodec.Read(after, outputContext);
            if (!MasterTextStylesSemanticHash(outputStyles).Equals(MasterTextStylesSemanticHash(requested.Masters[masterIndex].TextStyles), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_semantics_mismatch",
                    $"PPTX master {masterIndex + 1} text styles do not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData, outputContext)).Equals(BackgroundSemanticHash(requested.Masters[masterIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_master_background_mismatch",
                    $"PPTX master {masterIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputGraph[masterIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Masters[masterIndex].Placeholders,
                requested.Masters[masterIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputGraph[masterIndex].Part));
        }
        var sourceLayouts = sourceGraph.SelectMany(master => master.Layouts).ToArray();
        var outputLayouts = outputGraph.SelectMany(master => master.Layouts).ToArray();
        if (sourceLayouts.Length != requested.Layouts.Count || outputLayouts.Length != requested.Layouts.Count)
            throw new CodecException("presentation_postwrite_layout_topology_changed", "PPTX layout topology changed during source-preserving export.");
        for (var layoutIndex = 0; layoutIndex < requested.Layouts.Count; layoutIndex++)
        {
            var before = sourceLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX source layout {layoutIndex + 1} has no root.");
            var after = outputLayouts[layoutIndex].Part.SlideLayout ??
                throw new CodecException("missing_slide_layout_root", $"PPTX output layout {layoutIndex + 1} has no root.");
            var sourceContext = new PptxPartContext(sourceLayouts[layoutIndex].Part, sourceSlideMap, assets: sourceAssets, customShows: customShowCatalog);
            var outputContext = new PptxPartContext(outputLayouts[layoutIndex].Part, outputSlideMap, assets: outputAssets, customShows: customShowCatalog);
            if (!LayoutResidualHash(before, sourceContext).Equals(LayoutResidualHash(after, outputContext), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_unmodeled_layout_content_changed",
                    $"PPTX layout {layoutIndex + 1} edit changed unmodeled native content.",
                    PartPath(outputLayouts[layoutIndex].Part));
            if (!BackgroundSemanticHash(PptxBackgroundCodec.Read(after.CommonSlideData, outputContext)).Equals(BackgroundSemanticHash(requested.Layouts[layoutIndex].Background), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_postwrite_layout_background_mismatch",
                    $"PPTX layout {layoutIndex + 1} background does not match requested semantics after export.",
                    PartPath(outputLayouts[layoutIndex].Part));
            ValidatePlaceholderOutput(
                before.CommonSlideData?.ShapeTree,
                after.CommonSlideData?.ShapeTree,
                requested.Layouts[layoutIndex].Placeholders,
                requested.Layouts[layoutIndex].Id,
                sourceContext,
                outputContext,
                PartPath(outputLayouts[layoutIndex].Part));
        }
    }

    private static void ValidatePlaceholderOutput(
        P.ShapeTree? sourceTree,
        P.ShapeTree? outputTree,
        IList<PresentationPlaceholder> requested,
        string ownerId,
        PptxPartContext sourceContext,
        PptxPartContext outputContext,
        string partPath)
    {
        if (sourceTree is null || outputTree is null)
            throw new CodecException("missing_shape_tree", $"Presentation owner {ownerId} has no shape tree.", partPath);
        var before = PptxPlaceholderCodec.Read(sourceTree, ownerId, sourceContext);
        var after = PptxPlaceholderCodec.Read(outputTree, ownerId, outputContext);
        if (before.Count != requested.Count || after.Count != requested.Count)
            throw new CodecException("presentation_postwrite_placeholder_topology_changed", $"Presentation owner {ownerId} placeholder topology changed during export.", partPath);
        for (var index = 0; index < requested.Count; index++)
        {
            var request = requested[index];
            var binding = request.Source ?? throw new CodecException("missing_presentation_placeholder_binding", $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.", partPath);
            var changed = !PptxPlaceholderCodec.SemanticHash(request).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase);
            if (!changed)
            {
                var sourceShape = PptxPlaceholderCodec.BoundShape(sourceTree, before[index]);
                var outputShape = PptxPlaceholderCodec.BoundShape(outputTree, after[index]);
                if (sourceShape is null || outputShape is null ||
                    !PptxPlaceholderCodec.ElementHash(sourceShape).Equals(PptxPlaceholderCodec.ElementHash(outputShape), StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_unchanged_placeholder_modified", $"Presentation placeholder {index + 1} under {ownerId} was modified during export.", partPath);
            }
            if (after[index].Id != request.Id ||
                !PptxPlaceholderCodec.SemanticHash(after[index]).Equals(PptxPlaceholderCodec.SemanticHash(request), StringComparison.OrdinalIgnoreCase))
                throw new CodecException("presentation_postwrite_placeholder_semantics_mismatch", $"Presentation placeholder {index + 1} under {ownerId} does not match requested semantics after export.", partPath);
        }
    }

    private static SlidePart[] OrderedSlideParts(PresentationDocument package)
    {
        var presentationPart = package.PresentationPart ?? throw new CodecException("missing_presentation_part", "PPTX package has no Presentation part.", "ppt/presentation.xml");
        return ResolveSlideParts(presentationPart, presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? []);
    }

    private static PptxMasterGraphEntry[] ReadMasterGraph(PresentationPart presentationPart)
    {
        var masterIds = presentationPart.Presentation?.SlideMasterIdList?.Elements<P.SlideMasterId>().ToArray() ?? [];
        return masterIds.Select((masterId, masterIndex) =>
        {
            var relationshipId = masterId.RelationshipId?.Value ?? string.Empty;
            var masterPart = presentationPart.GetPartById(relationshipId) as SlideMasterPart ??
                throw new CodecException("missing_slide_master_part", $"Presentation master {masterIndex + 1} has an unresolved relationship.", "ppt/presentation.xml");
            var layoutIds = masterPart.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>().ToArray() ?? [];
            var layouts = layoutIds.Select((layoutId, layoutIndex) =>
            {
                var layoutRelationshipId = layoutId.RelationshipId?.Value ?? string.Empty;
                var layoutPart = masterPart.GetPartById(layoutRelationshipId) as SlideLayoutPart ??
                    throw new CodecException("missing_slide_layout_part", $"Presentation layout {layoutIndex + 1} under master {masterIndex + 1} has an unresolved relationship.", PartPath(masterPart));
                return new PptxLayoutGraphEntry(
                    layoutIndex,
                    $"presentation/master/{masterIndex + 1}/layout/{layoutIndex + 1}",
                    layoutRelationshipId,
                    layoutPart);
            }).ToArray();
            return new PptxMasterGraphEntry(
                masterIndex,
                $"presentation/master/{masterIndex + 1}",
                relationshipId,
                masterPart,
                layouts);
        }).ToArray();
    }

    private static ThemePart? CanonicalThemePart(IReadOnlyList<PptxMasterGraphEntry> masterGraph)
    {
        if (masterGraph.Count == 0 || masterGraph.Any(master => master.Part.ThemePart is null))
            return null;
        var themes = masterGraph
            .Select(master => master.Part.ThemePart!)
            .GroupBy(PartPath, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        return themes.Length == 1 ? themes[0] : null;
    }

    private static string[]? TryReadSourceBoundThemeAccentRgb(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme is null) return null;
        var colors = new[]
        {
            ReadSourceBoundThemeRgb(colorScheme.Accent1Color),
            ReadSourceBoundThemeRgb(colorScheme.Accent2Color),
            ReadSourceBoundThemeRgb(colorScheme.Accent3Color),
            ReadSourceBoundThemeRgb(colorScheme.Accent4Color),
            ReadSourceBoundThemeRgb(colorScheme.Accent5Color),
            ReadSourceBoundThemeRgb(colorScheme.Accent6Color),
        };
        return colors.Any(color => color is null) ? null : colors.Select(color => color!).ToArray();
    }

    private static uint? TryReadSourceBoundThemeAccent1Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent2Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent3Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent3Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent3LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent4HueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueModulation hueMod ||
            !HasOnlyAttributes(hueMod, "val") || hueMod.ChildElements.Count != 0 ||
            hueMod.Val is null || hueMod.Val.Value < 0 || hueMod.Val.Value > 100_000)
            return null;
        return checked((uint)hueMod.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent4LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent4LumOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceOffset lumOff ||
            !HasOnlyAttributes(lumOff, "val") || lumOff.ChildElements.Count != 0 ||
            lumOff.Val is null || lumOff.Val.Value < -100_000 || lumOff.Val.Value > 100_000)
            return null;
        return lumOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent4AlphaMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaModulation alphaMod ||
            !HasOnlyAttributes(alphaMod, "val") || alphaMod.ChildElements.Count != 0 ||
            alphaMod.Val is null || alphaMod.Val.Value < 0 || alphaMod.Val.Value > 100_000)
            return null;
        return checked((uint)alphaMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent4AlphaOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaOffset alphaOff ||
            !HasOnlyAttributes(alphaOff, "val") || alphaOff.ChildElements.Count != 0 ||
            alphaOff.Val is null || alphaOff.Val.Value < -100_000 || alphaOff.Val.Value > 100_000)
            return null;
        return alphaOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent4SatMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationModulation satMod ||
            !HasOnlyAttributes(satMod, "val") || satMod.ChildElements.Count != 0 ||
            satMod.Val is null || satMod.Val.Value < 0 || satMod.Val.Value > 100_000)
            return null;
        return checked((uint)satMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent4SatOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationOffset satOff ||
            !HasOnlyAttributes(satOff, "val") || satOff.ChildElements.Count != 0 ||
            satOff.Val is null || satOff.Val.Value < -100_000 || satOff.Val.Value > 100_000)
            return null;
        return satOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent4RedMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedModulation redMod ||
            !HasOnlyAttributes(redMod, "val") || redMod.ChildElements.Count != 0 ||
            redMod.Val is null || redMod.Val.Value < 0 || redMod.Val.Value > 100_000)
            return null;
        return checked((uint)redMod.Val.Value);
    }


    private static uint? TryReadSourceBoundThemeAccent4GreenMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenModulation greenMod ||
            !HasOnlyAttributes(greenMod, "val") || greenMod.ChildElements.Count != 0 ||
            greenMod.Val is null || greenMod.Val.Value < 0 || greenMod.Val.Value > 100_000)
            return null;
        return checked((uint)greenMod.Val.Value);
    }


    private static uint? TryReadSourceBoundThemeAccent4BlueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueModulation blueMod ||
            !HasOnlyAttributes(blueMod, "val") || blueMod.ChildElements.Count != 0 ||
            blueMod.Val is null || blueMod.Val.Value < 0 || blueMod.Val.Value > 100_000)
            return null;
        return checked((uint)blueMod.Val.Value);
    }


    private static int? TryReadSourceBoundThemeAccent4BlueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueOffset blueOff ||
            !HasOnlyAttributes(blueOff, "val") || blueOff.ChildElements.Count != 0 ||
            blueOff.Val is null || blueOff.Val.Value < -100_000 || blueOff.Val.Value > 100_000)
            return null;
        return blueOff.Val.Value;
    }


    private static int? TryReadSourceBoundThemeAccent4GreenOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenOffset greenOff ||
            !HasOnlyAttributes(greenOff, "val") || greenOff.ChildElements.Count != 0 ||
            greenOff.Val is null || greenOff.Val.Value < -100_000 || greenOff.Val.Value > 100_000)
            return null;
        return greenOff.Val.Value;
    }


    private static int? TryReadSourceBoundThemeAccent4RedOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedOffset redOff ||
            !HasOnlyAttributes(redOff, "val") || redOff.ChildElements.Count != 0 ||
            redOff.Val is null || redOff.Val.Value < -100_000 || redOff.Val.Value > 100_000)
            return null;
        return redOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent3AlphaMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaModulation alphaMod ||
            !HasOnlyAttributes(alphaMod, "val") || alphaMod.ChildElements.Count != 0 ||
            alphaMod.Val is null || alphaMod.Val.Value < 0 || alphaMod.Val.Value > 100_000)
            return null;
        return checked((uint)alphaMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent3AlphaOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaOffset alphaOff ||
            !HasOnlyAttributes(alphaOff, "val") || alphaOff.ChildElements.Count != 0 ||
            alphaOff.Val is null || alphaOff.Val.Value < -100_000 || alphaOff.Val.Value > 100_000)
            return null;
        return alphaOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent3SatMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationModulation satMod ||
            !HasOnlyAttributes(satMod, "val") || satMod.ChildElements.Count != 0 ||
            satMod.Val is null || satMod.Val.Value < 0 || satMod.Val.Value > 100_000)
            return null;
        return checked((uint)satMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent3SatOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationOffset satOff ||
            !HasOnlyAttributes(satOff, "val") || satOff.ChildElements.Count != 0 ||
            satOff.Val is null || satOff.Val.Value < -100_000 || satOff.Val.Value > 100_000)
            return null;
        return satOff.Val.Value;
    }


    private static uint? TryReadSourceBoundThemeAccent3RedMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedModulation redMod ||
            !HasOnlyAttributes(redMod, "val") || redMod.ChildElements.Count != 0 ||
            redMod.Val is null || redMod.Val.Value < 0 || redMod.Val.Value > 100_000)
            return null;
        return checked((uint)redMod.Val.Value);
    }


    private static int? TryReadSourceBoundThemeAccent3RedOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedOffset redOff ||
            !HasOnlyAttributes(redOff, "val") || redOff.ChildElements.Count != 0 ||
            redOff.Val is null || redOff.Val.Value < -100_000 || redOff.Val.Value > 100_000)
            return null;
        return redOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent3GreenMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenModulation greenMod ||
            !HasOnlyAttributes(greenMod, "val") || greenMod.ChildElements.Count != 0 ||
            greenMod.Val is null || greenMod.Val.Value < 0 || greenMod.Val.Value > 100_000)
            return null;
        return checked((uint)greenMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent3GreenOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenOffset greenOff ||
            !HasOnlyAttributes(greenOff, "val") || greenOff.ChildElements.Count != 0 ||
            greenOff.Val is null || greenOff.Val.Value < -100_000 || greenOff.Val.Value > 100_000)
            return null;
        return greenOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent3BlueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueModulation blueMod ||
            !HasOnlyAttributes(blueMod, "val") || blueMod.ChildElements.Count != 0 ||
            blueMod.Val is null || blueMod.Val.Value < 0 || blueMod.Val.Value > 100_000)
            return null;
        return checked((uint)blueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent3BlueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueOffset blueOff ||
            !HasOnlyAttributes(blueOff, "val") || blueOff.ChildElements.Count != 0 ||
            blueOff.Val is null || blueOff.Val.Value < -100_000 || blueOff.Val.Value > 100_000)
            return null;
        return blueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent3HueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueModulation hueMod ||
            !HasOnlyAttributes(hueMod, "val") || hueMod.ChildElements.Count != 0 ||
            hueMod.Val is null || hueMod.Val.Value < 0 || hueMod.Val.Value > 100_000)
            return null;
        return checked((uint)hueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent3HueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueOffset hueOff ||
            !HasOnlyAttributes(hueOff, "val") || hueOff.ChildElements.Count != 0 ||
            hueOff.Val is null || hueOff.Val.Value < -21_600_000 || hueOff.Val.Value > 21_600_000)
            return null;
        return hueOff.Val.Value;
    }

    private static int? TryReadSourceBoundThemeAccent4HueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueOffset hueOff ||
            !HasOnlyAttributes(hueOff, "val") || hueOff.ChildElements.Count != 0 ||
            hueOff.Val is null || hueOff.Val.Value < -21_600_000 || hueOff.Val.Value > 21_600_000)
            return null;
        return hueOff.Val.Value;
    }

    private static int? TryReadSourceBoundThemeAccent3LumOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent3Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceOffset lumOff ||
            !HasOnlyAttributes(lumOff, "val") || lumOff.ChildElements.Count != 0 ||
            lumOff.Val is null || lumOff.Val.Value < -100_000 || lumOff.Val.Value > 100_000)
            return null;
        return checked(lumOff.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent2Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent2LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2LumOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceOffset lumOff ||
            !HasOnlyAttributes(lumOff, "val") || lumOff.ChildElements.Count != 0 ||
            lumOff.Val is null || lumOff.Val.Value < -100_000 || lumOff.Val.Value > 100_000)
            return null;
        return lumOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent2AlphaMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaModulation alphaMod ||
            !HasOnlyAttributes(alphaMod, "val") || alphaMod.ChildElements.Count != 0 ||
            alphaMod.Val is null || alphaMod.Val.Value < 0 || alphaMod.Val.Value > 100_000)
            return null;
        return checked((uint)alphaMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2AlphaOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaOffset alphaOff ||
            !HasOnlyAttributes(alphaOff, "val") || alphaOff.ChildElements.Count != 0 ||
            alphaOff.Val is null || alphaOff.Val.Value < -100_000 || alphaOff.Val.Value > 100_000)
            return null;
        return alphaOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent2SatMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationModulation satMod ||
            !HasOnlyAttributes(satMod, "val") || satMod.ChildElements.Count != 0 ||
            satMod.Val is null || satMod.Val.Value < 0 || satMod.Val.Value > 100_000)
            return null;
        return checked((uint)satMod.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent2RedMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedModulation redMod ||
            !HasOnlyAttributes(redMod, "val") || redMod.ChildElements.Count != 0 ||
            redMod.Val is null || redMod.Val.Value < 0 || redMod.Val.Value > 100_000)
            return null;
        return checked((uint)redMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2SatOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationOffset satOff ||
            !HasOnlyAttributes(satOff, "val") || satOff.ChildElements.Count != 0 ||
            satOff.Val is null || satOff.Val.Value < -100_000 || satOff.Val.Value > 100_000)
            return null;
        return satOff.Val.Value;
    }

    private static int? TryReadSourceBoundThemeAccent2RedOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedOffset redOff ||
            !HasOnlyAttributes(redOff, "val") || redOff.ChildElements.Count != 0 ||
            redOff.Val is null || redOff.Val.Value < -100_000 || redOff.Val.Value > 100_000)
            return null;
        return redOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent2GreenMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenModulation greenMod ||
            !HasOnlyAttributes(greenMod, "val") || greenMod.ChildElements.Count != 0 ||
            greenMod.Val is null || greenMod.Val.Value < 0 || greenMod.Val.Value > 100_000)
            return null;
        return checked((uint)greenMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2GreenOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenOffset greenOff ||
            !HasOnlyAttributes(greenOff, "val") || greenOff.ChildElements.Count != 0 ||
            greenOff.Val is null || greenOff.Val.Value < -100_000 || greenOff.Val.Value > 100_000)
            return null;
        return greenOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent2BlueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueModulation blueMod ||
            !HasOnlyAttributes(blueMod, "val") || blueMod.ChildElements.Count != 0 ||
            blueMod.Val is null || blueMod.Val.Value < 0 || blueMod.Val.Value > 100_000)
            return null;
        return checked((uint)blueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2BlueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueOffset blueOff ||
            !HasOnlyAttributes(blueOff, "val") || blueOff.ChildElements.Count != 0 ||
            blueOff.Val is null || blueOff.Val.Value < -100_000 || blueOff.Val.Value > 100_000)
            return null;
        return blueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent2HueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueModulation hueMod ||
            !HasOnlyAttributes(hueMod, "val") || hueMod.ChildElements.Count != 0 ||
            hueMod.Val is null || hueMod.Val.Value < 0 || hueMod.Val.Value > 100_000)
            return null;
        return checked((uint)hueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent2HueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent2Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueOffset hueOff ||
            !HasOnlyAttributes(hueOff, "val") || hueOff.ChildElements.Count != 0 ||
            hueOff.Val is null || hueOff.Val.Value < -21_600_000 || hueOff.Val.Value > 21_600_000)
            return null;
        return hueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent1LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1LumOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceOffset lumOff ||
            !HasOnlyAttributes(lumOff, "val") || lumOff.ChildElements.Count != 0 ||
            lumOff.Val is null || lumOff.Val.Value < -100_000 || lumOff.Val.Value > 100_000)
            return null;
        return lumOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1AlphaMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaModulation alphaMod ||
            !HasOnlyAttributes(alphaMod, "val") || alphaMod.ChildElements.Count != 0 ||
            alphaMod.Val is null || alphaMod.Val.Value < 0 || alphaMod.Val.Value > 100_000)
            return null;
        return checked((uint)alphaMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1AlphaOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaOffset alphaOff ||
            !HasOnlyAttributes(alphaOff, "val") || alphaOff.ChildElements.Count != 0 ||
            alphaOff.Val is null || alphaOff.Val.Value < -100_000 || alphaOff.Val.Value > 100_000)
            return null;
        return alphaOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1SatMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationModulation satMod ||
            !HasOnlyAttributes(satMod, "val") || satMod.ChildElements.Count != 0 ||
            satMod.Val is null || satMod.Val.Value < 0 || satMod.Val.Value > 100_000)
            return null;
        return checked((uint)satMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1SatOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationOffset satOff ||
            !HasOnlyAttributes(satOff, "val") || satOff.ChildElements.Count != 0 ||
            satOff.Val is null || satOff.Val.Value < -100_000 || satOff.Val.Value > 100_000)
            return null;
        return satOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1RedMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedModulation redMod ||
            !HasOnlyAttributes(redMod, "val") || redMod.ChildElements.Count != 0 ||
            redMod.Val is null || redMod.Val.Value < 0 || redMod.Val.Value > 100_000)
            return null;
        return checked((uint)redMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1RedOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedOffset redOff ||
            !HasOnlyAttributes(redOff, "val") || redOff.ChildElements.Count != 0 ||
            redOff.Val is null || redOff.Val.Value < -100_000 || redOff.Val.Value > 100_000)
            return null;
        return redOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1GreenMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenModulation greenMod ||
            !HasOnlyAttributes(greenMod, "val") || greenMod.ChildElements.Count != 0 ||
            greenMod.Val is null || greenMod.Val.Value < 0 || greenMod.Val.Value > 100_000)
            return null;
        return checked((uint)greenMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1GreenOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenOffset greenOff ||
            !HasOnlyAttributes(greenOff, "val") || greenOff.ChildElements.Count != 0 ||
            greenOff.Val is null || greenOff.Val.Value < -100_000 || greenOff.Val.Value > 100_000)
            return null;
        return greenOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1BlueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueModulation blueMod ||
            !HasOnlyAttributes(blueMod, "val") || blueMod.ChildElements.Count != 0 ||
            blueMod.Val is null || blueMod.Val.Value < 0 || blueMod.Val.Value > 100_000)
            return null;
        return checked((uint)blueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1BlueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueOffset blueOff ||
            !HasOnlyAttributes(blueOff, "val") || blueOff.ChildElements.Count != 0 ||
            blueOff.Val is null || blueOff.Val.Value < -100_000 || blueOff.Val.Value > 100_000)
            return null;
        return blueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent1HueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueModulation hueMod ||
            !HasOnlyAttributes(hueMod, "val") || hueMod.ChildElements.Count != 0 ||
            hueMod.Val is null || hueMod.Val.Value < 0 || hueMod.Val.Value > 100_000)
            return null;
        return checked((uint)hueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent1HueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent1Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueOffset hueOff ||
            !HasOnlyAttributes(hueOff, "val") || hueOff.ChildElements.Count != 0 ||
            hueOff.Val is null || hueOff.Val.Value < -21_600_000 || hueOff.Val.Value > 21_600_000)
            return null;
        return hueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent4Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent4Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent4Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent5Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent5Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent5LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent6LumMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent6Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceModulation lumMod ||
            !HasOnlyAttributes(lumMod, "val") || lumMod.ChildElements.Count != 0 ||
            lumMod.Val is null || lumMod.Val.Value < 0 || lumMod.Val.Value > 100_000)
            return null;
        return checked((uint)lumMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5LumOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.LuminanceOffset lumOff ||
            !HasOnlyAttributes(lumOff, "val") || lumOff.ChildElements.Count != 0 ||
            lumOff.Val is null || lumOff.Val.Value < -100_000 || lumOff.Val.Value > 100_000)
            return null;
        return lumOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5AlphaMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaModulation alphaMod ||
            !HasOnlyAttributes(alphaMod, "val") || alphaMod.ChildElements.Count != 0 ||
            alphaMod.Val is null || alphaMod.Val.Value < 0 || alphaMod.Val.Value > 100_000)
            return null;
        return checked((uint)alphaMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5AlphaOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.AlphaOffset alphaOff ||
            !HasOnlyAttributes(alphaOff, "val") || alphaOff.ChildElements.Count != 0 ||
            alphaOff.Val is null || alphaOff.Val.Value < -100_000 || alphaOff.Val.Value > 100_000)
            return null;
        return alphaOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5SatMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationModulation satMod ||
            !HasOnlyAttributes(satMod, "val") || satMod.ChildElements.Count != 0 ||
            satMod.Val is null || satMod.Val.Value < 0 || satMod.Val.Value > 100_000)
            return null;
        return checked((uint)satMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5SatOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.SaturationOffset satOff ||
            !HasOnlyAttributes(satOff, "val") || satOff.ChildElements.Count != 0 ||
            satOff.Val is null || satOff.Val.Value < -100_000 || satOff.Val.Value > 100_000)
            return null;
        return satOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5RedMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedModulation redMod ||
            !HasOnlyAttributes(redMod, "val") || redMod.ChildElements.Count != 0 ||
            redMod.Val is null || redMod.Val.Value < 0 || redMod.Val.Value > 100_000)
            return null;
        return checked((uint)redMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5RedOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.RedOffset redOff ||
            !HasOnlyAttributes(redOff, "val") || redOff.ChildElements.Count != 0 ||
            redOff.Val is null || redOff.Val.Value < -100_000 || redOff.Val.Value > 100_000)
            return null;
        return redOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5GreenMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenModulation greenMod ||
            !HasOnlyAttributes(greenMod, "val") || greenMod.ChildElements.Count != 0 ||
            greenMod.Val is null || greenMod.Val.Value < 0 || greenMod.Val.Value > 100_000)
            return null;
        return checked((uint)greenMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5GreenOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.GreenOffset greenOff ||
            !HasOnlyAttributes(greenOff, "val") || greenOff.ChildElements.Count != 0 ||
            greenOff.Val is null || greenOff.Val.Value < -100_000 || greenOff.Val.Value > 100_000)
            return null;
        return greenOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5BlueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueModulation blueMod ||
            !HasOnlyAttributes(blueMod, "val") || blueMod.ChildElements.Count != 0 ||
            blueMod.Val is null || blueMod.Val.Value < 0 || blueMod.Val.Value > 100_000)
            return null;
        return checked((uint)blueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5BlueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.BlueOffset blueOff ||
            !HasOnlyAttributes(blueOff, "val") || blueOff.ChildElements.Count != 0 ||
            blueOff.Val is null || blueOff.Val.Value < -100_000 || blueOff.Val.Value > 100_000)
            return null;
        return blueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent5HueMod(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueModulation hueMod ||
            !HasOnlyAttributes(hueMod, "val") || hueMod.ChildElements.Count != 0 ||
            hueMod.Val is null || hueMod.Val.Value < 0 || hueMod.Val.Value > 100_000)
            return null;
        return checked((uint)hueMod.Val.Value);
    }

    private static int? TryReadSourceBoundThemeAccent5HueOff(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent5Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.HueOffset hueOff ||
            !HasOnlyAttributes(hueOff, "val") || hueOff.ChildElements.Count != 0 ||
            hueOff.Val is null || hueOff.Val.Value < -21_600_000 || hueOff.Val.Value > 21_600_000)
            return null;
        return hueOff.Val.Value;
    }

    private static uint? TryReadSourceBoundThemeAccent6Tint(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent6Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Tint tint ||
            !HasOnlyAttributes(tint, "val") || tint.ChildElements.Count != 0 ||
            tint.Val is null || tint.Val.Value < 0 || tint.Val.Value > 100_000)
            return null;
        return checked((uint)tint.Val.Value);
    }

    private static uint? TryReadSourceBoundThemeAccent6Shade(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        var owner = colorScheme?.Accent6Color;
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) ||
            color.ChildElements.Count != 1 || color.FirstChild is not A.Shade shade ||
            !HasOnlyAttributes(shade, "val") || shade.ChildElements.Count != 0 ||
            shade.Val is null || shade.Val.Value < 0 || shade.Val.Value > 100_000)
            return null;
        return checked((uint)shade.Val.Value);
    }

    private static string[]? TryReadSourceBoundThemeColorRoleRgb(A.Theme? theme)
    {
        var colorScheme = theme?.ThemeElements?.ColorScheme;
        if (colorScheme is null) return null;
        var colors = new[]
        {
            ReadSourceBoundThemeRgb(colorScheme.Dark1Color),
            ReadSourceBoundThemeRgb(colorScheme.Light1Color),
            ReadSourceBoundThemeRgb(colorScheme.Dark2Color),
            ReadSourceBoundThemeRgb(colorScheme.Light2Color),
            ReadSourceBoundThemeRgb(colorScheme.Hyperlink),
            ReadSourceBoundThemeRgb(colorScheme.FollowedHyperlinkColor),
        };
        return colors.Any(color => color is null) ? null : colors.Select(color => color!).ToArray();
    }

    private static string? ReadSourceBoundThemeRgb(OpenXmlElement? owner)
    {
        if (owner is null || !HasOnlyAttributes(owner) || owner.ChildElements.Count != 1 ||
            owner.FirstChild is not A.RgbColorModelHex color ||
            !HasOnlyAttributes(color, "val") || color.ChildElements.Count != 0 ||
            color.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit))
            return null;
        return value.ToUpperInvariant();
    }

    private static void ApplySourceBoundThemeFields(
        PresentationThemeArtifact? authoredTheme,
        IReadOnlyList<PptxMasterGraphEntry> masterGraph,
        ISet<string> changedParts,
        IDictionary<string, string> replacedOpaquePartHashes)
    {
        if (authoredTheme is null ||
            (authoredTheme.AccentRgb.Count == 0 &&
             authoredTheme.AccentTransforms.Count == 0 &&
             !authoredTheme.HasDark1Rgb && !authoredTheme.HasLight1Rgb && !authoredTheme.HasDark2Rgb &&
             !authoredTheme.HasLight2Rgb && !authoredTheme.HasHyperlinkRgb && !authoredTheme.HasFollowedHyperlinkRgb &&
             !authoredTheme.HasName && !authoredTheme.HasMajorFontFamily && !authoredTheme.HasMinorFontFamily &&
             !authoredTheme.HasMajorFontFamilyEastAsia &&
             !authoredTheme.HasMinorFontFamilyEastAsia && !authoredTheme.HasMajorFontFamilyComplexScript &&
             !authoredTheme.HasMinorFontFamilyComplexScript)) return;
        var themePart = CanonicalThemePart(masterGraph) ??
            throw new CodecException(
                "unsupported_presentation_edit",
                "Source-preserving PPTX export cannot edit a theme without one shared ThemePart.",
                "ppt/presentation.xml");
        var theme = themePart.Theme ??
            throw new CodecException(
                "missing_theme_root",
                "PPTX source has no theme root.",
                PartPath(themePart));
        var changed = false;
        if (authoredTheme.HasName)
        {
            var sourceName = theme.Name?.Value;
            if (string.IsNullOrWhiteSpace(sourceName))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing theme name.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.Name))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove a theme name in this bounded profile.",
                    "$.design.theme.name");
            if (!sourceName.Equals(authoredTheme.Name, StringComparison.Ordinal))
            {
                theme.Name = authoredTheme.Name;
                changed = true;
            }
        }
        if (authoredTheme.HasMajorFontFamily)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMajor = fontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMajor))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing major Latin theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MajorFontFamily))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the major Latin theme font in this bounded profile.",
                    "$.design.theme.fontScheme.major");
            if (!sourceMajor.Equals(authoredTheme.MajorFontFamily, StringComparison.Ordinal))
            {
                if (fontScheme?.MajorFont?.LatinFont is not { } latinFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing major Latin theme font node.",
                        PartPath(themePart));
                latinFont.Typeface = authoredTheme.MajorFontFamily;
                changed = true;
            }
        }
        if (authoredTheme.HasMinorFontFamily)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMinor = fontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMinor))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing minor Latin theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MinorFontFamily))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the minor Latin theme font in this bounded profile.",
                    "$.design.theme.fontScheme.minor");
            if (!sourceMinor.Equals(authoredTheme.MinorFontFamily, StringComparison.Ordinal))
            {
                if (fontScheme?.MinorFont?.LatinFont is not { } latinFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing minor Latin theme font node.",
                        PartPath(themePart));
                latinFont.Typeface = authoredTheme.MinorFontFamily;
                changed = true;
            }
        }
        if (authoredTheme.HasMajorFontFamilyEastAsia)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMajorEastAsia = fontScheme?.MajorFont?.EastAsianFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMajorEastAsia))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing major East Asian theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MajorFontFamilyEastAsia))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the major East Asian theme font in this bounded profile.",
                    "$.design.theme.fontScheme.majorEastAsia");
            if (!sourceMajorEastAsia.Equals(authoredTheme.MajorFontFamilyEastAsia, StringComparison.Ordinal))
            {
                if (fontScheme?.MajorFont?.EastAsianFont is not { } eastAsianFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing major East Asian theme font node.",
                        PartPath(themePart));
                eastAsianFont.Typeface = authoredTheme.MajorFontFamilyEastAsia;
                changed = true;
            }
        }
        if (authoredTheme.HasMinorFontFamilyEastAsia)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMinorEastAsia = fontScheme?.MinorFont?.EastAsianFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMinorEastAsia))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing minor East Asian theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MinorFontFamilyEastAsia))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the minor East Asian theme font in this bounded profile.",
                    "$.design.theme.fontScheme.minorEastAsia");
            if (!sourceMinorEastAsia.Equals(authoredTheme.MinorFontFamilyEastAsia, StringComparison.Ordinal))
            {
                if (fontScheme?.MinorFont?.EastAsianFont is not { } eastAsianFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing minor East Asian theme font node.",
                        PartPath(themePart));
                eastAsianFont.Typeface = authoredTheme.MinorFontFamilyEastAsia;
                changed = true;
            }
        }
        if (authoredTheme.HasMajorFontFamilyComplexScript)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMajorComplexScript = fontScheme?.MajorFont?.ComplexScriptFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMajorComplexScript))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing major complex-script theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MajorFontFamilyComplexScript))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the major complex-script theme font in this bounded profile.",
                    "$.design.theme.fontScheme.majorComplexScript");
            if (!sourceMajorComplexScript.Equals(authoredTheme.MajorFontFamilyComplexScript, StringComparison.Ordinal))
            {
                if (fontScheme?.MajorFont?.ComplexScriptFont is not { } complexScriptFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing major complex-script theme font node.",
                        PartPath(themePart));
                complexScriptFont.Typeface = authoredTheme.MajorFontFamilyComplexScript;
                changed = true;
            }
        }
        if (authoredTheme.HasMinorFontFamilyComplexScript)
        {
            var fontScheme = theme.ThemeElements?.FontScheme;
            var sourceMinorComplexScript = fontScheme?.MinorFont?.ComplexScriptFont?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(sourceMinorComplexScript))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing minor complex-script theme font.",
                    PartPath(themePart));
            if (string.IsNullOrWhiteSpace(authoredTheme.MinorFontFamilyComplexScript))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export cannot remove the minor complex-script theme font in this bounded profile.",
                    "$.design.theme.fontScheme.minorComplexScript");
            if (!sourceMinorComplexScript.Equals(authoredTheme.MinorFontFamilyComplexScript, StringComparison.Ordinal))
            {
                if (fontScheme?.MinorFont?.ComplexScriptFont is not { } complexScriptFont)
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing minor complex-script theme font node.",
                        PartPath(themePart));
                complexScriptFont.Typeface = authoredTheme.MinorFontFamilyComplexScript;
                changed = true;
            }
        }
        if (authoredTheme.HasDark1Rgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceDark1 = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.Dark1Color);
            if (sourceDark1 is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB dark1 color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.Dark1Rgb);
            if (!sourceDark1.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.Dark1Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct dark1 color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.HasLight1Rgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceLight1 = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.Light1Color);
            if (sourceLight1 is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB light1 color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.Light1Rgb);
            if (!sourceLight1.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.Light1Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct light1 color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.HasDark2Rgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceDark2 = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.Dark2Color);
            if (sourceDark2 is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB dark2 color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.Dark2Rgb);
            if (!sourceDark2.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.Dark2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct dark2 color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.HasLight2Rgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceLight2 = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.Light2Color);
            if (sourceLight2 is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB light2 color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.Light2Rgb);
            if (!sourceLight2.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.Light2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct light2 color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.HasHyperlinkRgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceHyperlink = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.Hyperlink);
            if (sourceHyperlink is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB hyperlink color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.HyperlinkRgb);
            if (!sourceHyperlink.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.Hyperlink?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct hyperlink color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.HasFollowedHyperlinkRgb)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceFollowedHyperlink = colorScheme is null
                ? null
                : ReadSourceBoundThemeRgb(colorScheme.FollowedHyperlinkColor);
            if (sourceFollowedHyperlink is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only a direct RGB followed-hyperlink color.",
                    PartPath(themePart));
            var requested = PptxColor.Normalize(authoredTheme.FollowedHyperlinkRgb);
            if (!sourceFollowedHyperlink.Equals(requested, StringComparison.OrdinalIgnoreCase))
            {
                var color = colorScheme?.FollowedHyperlinkColor?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct followed-hyperlink color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.AccentRgb.Count > 0)
        {
            if (authoredTheme.AccentRgb.Count != 6)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export requires all six imported accent colors.",
                    "$.design.theme.accentColors");
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceAccents = colorScheme is null
                ? null
                : new[]
                {
                    ReadSourceBoundThemeRgb(colorScheme.Accent1Color),
                    ReadSourceBoundThemeRgb(colorScheme.Accent2Color),
                    ReadSourceBoundThemeRgb(colorScheme.Accent3Color),
                    ReadSourceBoundThemeRgb(colorScheme.Accent4Color),
                    ReadSourceBoundThemeRgb(colorScheme.Accent5Color),
                    ReadSourceBoundThemeRgb(colorScheme.Accent6Color),
                };
            if (sourceAccents is null || sourceAccents.Any(color => color is null))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only six direct RGB accent colors.",
                    PartPath(themePart));
            OpenXmlElement[] accentOwners =
            {
                colorScheme!.Accent1Color!,
                colorScheme.Accent2Color!,
                colorScheme.Accent3Color!,
                colorScheme.Accent4Color!,
                colorScheme.Accent5Color!,
                colorScheme.Accent6Color!,
            };
            for (var index = 0; index < accentOwners.Length; index++)
            {
                var requested = PptxColor.Normalize(authoredTheme.AccentRgb[index]);
                if (sourceAccents[index]!.Equals(requested, StringComparison.OrdinalIgnoreCase))
                    continue;
                var color = accentOwners[index].GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent color.",
                        PartPath(themePart));
                color.Val = requested;
                changed = true;
            }
        }
        if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasTintThousandth: true } accent2Tint &&
            !accent2Tint.HasShadeThousandth &&
            !accent2Tint.HasLuminanceModulationThousandth &&
            !accent2Tint.HasLuminanceOffsetThousandth &&
            !accent2Tint.HasAlphaModulationThousandth &&
            !accent2Tint.HasAlphaOffsetThousandth &&
            !accent2Tint.HasSaturationModulationThousandth &&
            !accent2Tint.HasSaturationOffsetThousandth &&
            !accent2Tint.HasRedModulationThousandth &&
            !accent2Tint.HasRedOffsetThousandth &&
            !accent2Tint.HasGreenModulationThousandth &&
            !accent2Tint.HasGreenOffsetThousandth &&
            !accent2Tint.HasBlueModulationThousandth &&
            !accent2Tint.HasBlueOffsetThousandth &&
            !accent2Tint.HasHueModulationThousandth &&
            !accent2Tint.HasHueOffsetAngleThousandth &&
            !accent2Tint.HasGray && !accent2Tint.HasComp && !accent2Tint.HasInv &&
            !accent2Tint.HasGamma && !accent2Tint.HasInvGamma &&
            accent2Tint.TintThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2Tint(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 tint.",
                    PartPath(themePart));
            var requested = (long)accent2Tint.TintThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var tint = color.FirstChild as A.Tint ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 tint.",
                        PartPath(themePart));
                tint.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasShadeThousandth: true } accent2Shade &&
            !accent2Shade.HasTintThousandth &&
            !accent2Shade.HasLuminanceModulationThousandth &&
            !accent2Shade.HasLuminanceOffsetThousandth &&
            !accent2Shade.HasAlphaModulationThousandth &&
            !accent2Shade.HasAlphaOffsetThousandth &&
            !accent2Shade.HasSaturationModulationThousandth &&
            !accent2Shade.HasSaturationOffsetThousandth &&
            !accent2Shade.HasRedModulationThousandth &&
            !accent2Shade.HasRedOffsetThousandth &&
            !accent2Shade.HasGreenModulationThousandth &&
            !accent2Shade.HasGreenOffsetThousandth &&
            !accent2Shade.HasBlueModulationThousandth &&
            !accent2Shade.HasBlueOffsetThousandth &&
            !accent2Shade.HasHueModulationThousandth &&
            !accent2Shade.HasHueOffsetAngleThousandth &&
            !accent2Shade.HasGray && !accent2Shade.HasComp && !accent2Shade.HasInv &&
            !accent2Shade.HasGamma && !accent2Shade.HasInvGamma &&
            accent2Shade.ShadeThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2Shade(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 shade.",
                    PartPath(themePart));
            var requested = (long)accent2Shade.ShadeThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var shade = color.FirstChild as A.Shade ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 shade.",
                        PartPath(themePart));
                shade.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasLuminanceModulationThousandth: true } accent2LumMod &&
            !accent2LumMod.HasTintThousandth &&
            !accent2LumMod.HasShadeThousandth &&
            !accent2LumMod.HasLuminanceOffsetThousandth &&
            !accent2LumMod.HasAlphaModulationThousandth &&
            !accent2LumMod.HasAlphaOffsetThousandth &&
            !accent2LumMod.HasSaturationModulationThousandth &&
            !accent2LumMod.HasSaturationOffsetThousandth &&
            !accent2LumMod.HasRedModulationThousandth &&
            !accent2LumMod.HasRedOffsetThousandth &&
            !accent2LumMod.HasGreenModulationThousandth &&
            !accent2LumMod.HasGreenOffsetThousandth &&
            !accent2LumMod.HasBlueModulationThousandth &&
            !accent2LumMod.HasBlueOffsetThousandth &&
            !accent2LumMod.HasHueModulationThousandth &&
            !accent2LumMod.HasHueOffsetAngleThousandth &&
            !accent2LumMod.HasGray && !accent2LumMod.HasComp && !accent2LumMod.HasInv &&
            !accent2LumMod.HasGamma && !accent2LumMod.HasInvGamma &&
            accent2LumMod.LuminanceModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2LumMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 lumMod.",
                    PartPath(themePart));
            var requested = (long)accent2LumMod.LuminanceModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var lumMod = color.FirstChild as A.LuminanceModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 lumMod.",
                        PartPath(themePart));
                lumMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasLuminanceOffsetThousandth: true } accent2LumOff &&
            !accent2LumOff.HasTintThousandth &&
            !accent2LumOff.HasShadeThousandth &&
            !accent2LumOff.HasLuminanceModulationThousandth &&
            !accent2LumOff.HasAlphaModulationThousandth &&
            !accent2LumOff.HasAlphaOffsetThousandth &&
            !accent2LumOff.HasSaturationModulationThousandth &&
            !accent2LumOff.HasSaturationOffsetThousandth &&
            !accent2LumOff.HasRedModulationThousandth &&
            !accent2LumOff.HasRedOffsetThousandth &&
            !accent2LumOff.HasGreenModulationThousandth &&
            !accent2LumOff.HasGreenOffsetThousandth &&
            !accent2LumOff.HasBlueModulationThousandth &&
            !accent2LumOff.HasBlueOffsetThousandth &&
            !accent2LumOff.HasHueModulationThousandth &&
            !accent2LumOff.HasHueOffsetAngleThousandth &&
            !accent2LumOff.HasGray && !accent2LumOff.HasComp && !accent2LumOff.HasInv &&
            !accent2LumOff.HasGamma && !accent2LumOff.HasInvGamma &&
            accent2LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent2LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2LumOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 lumOff.",
                    PartPath(themePart));
            var requested = (long)accent2LumOff.LuminanceOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var lumOff = color.FirstChild as A.LuminanceOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 lumOff.",
                        PartPath(themePart));
                lumOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasAlphaModulationThousandth: true } accent2AlphaMod &&
            !accent2AlphaMod.HasTintThousandth &&
            !accent2AlphaMod.HasShadeThousandth &&
            !accent2AlphaMod.HasLuminanceModulationThousandth &&
            !accent2AlphaMod.HasLuminanceOffsetThousandth &&
            !accent2AlphaMod.HasAlphaOffsetThousandth &&
            !accent2AlphaMod.HasSaturationModulationThousandth &&
            !accent2AlphaMod.HasSaturationOffsetThousandth &&
            !accent2AlphaMod.HasRedModulationThousandth &&
            !accent2AlphaMod.HasRedOffsetThousandth &&
            !accent2AlphaMod.HasGreenModulationThousandth &&
            !accent2AlphaMod.HasGreenOffsetThousandth &&
            !accent2AlphaMod.HasBlueModulationThousandth &&
            !accent2AlphaMod.HasBlueOffsetThousandth &&
            !accent2AlphaMod.HasHueModulationThousandth &&
            !accent2AlphaMod.HasHueOffsetAngleThousandth &&
            !accent2AlphaMod.HasGray && !accent2AlphaMod.HasComp && !accent2AlphaMod.HasInv &&
            !accent2AlphaMod.HasGamma && !accent2AlphaMod.HasInvGamma &&
            accent2AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2AlphaMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 alphaMod.",
                    PartPath(themePart));
            var requested = (long)accent2AlphaMod.AlphaModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var alphaMod = color.FirstChild as A.AlphaModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 alphaMod.",
                        PartPath(themePart));
                alphaMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasAlphaOffsetThousandth: true } accent2AlphaOff &&
            !accent2AlphaOff.HasTintThousandth &&
            !accent2AlphaOff.HasShadeThousandth &&
            !accent2AlphaOff.HasLuminanceModulationThousandth &&
            !accent2AlphaOff.HasLuminanceOffsetThousandth &&
            !accent2AlphaOff.HasAlphaModulationThousandth &&
            !accent2AlphaOff.HasSaturationModulationThousandth &&
            !accent2AlphaOff.HasSaturationOffsetThousandth &&
            !accent2AlphaOff.HasRedModulationThousandth &&
            !accent2AlphaOff.HasRedOffsetThousandth &&
            !accent2AlphaOff.HasGreenModulationThousandth &&
            !accent2AlphaOff.HasGreenOffsetThousandth &&
            !accent2AlphaOff.HasBlueModulationThousandth &&
            !accent2AlphaOff.HasBlueOffsetThousandth &&
            !accent2AlphaOff.HasHueModulationThousandth &&
            !accent2AlphaOff.HasHueOffsetAngleThousandth &&
            !accent2AlphaOff.HasGray && !accent2AlphaOff.HasComp && !accent2AlphaOff.HasInv &&
            !accent2AlphaOff.HasGamma && !accent2AlphaOff.HasInvGamma &&
            accent2AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent2AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2AlphaOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 alphaOff.",
                    PartPath(themePart));
            var requested = (long)accent2AlphaOff.AlphaOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var alphaOff = color.FirstChild as A.AlphaOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 alphaOff.",
                        PartPath(themePart));
                alphaOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasSaturationModulationThousandth: true } accent2SatMod &&
            !accent2SatMod.HasTintThousandth &&
            !accent2SatMod.HasShadeThousandth &&
            !accent2SatMod.HasLuminanceModulationThousandth &&
            !accent2SatMod.HasLuminanceOffsetThousandth &&
            !accent2SatMod.HasAlphaModulationThousandth &&
            !accent2SatMod.HasAlphaOffsetThousandth &&
            !accent2SatMod.HasSaturationOffsetThousandth &&
            !accent2SatMod.HasRedModulationThousandth &&
            !accent2SatMod.HasRedOffsetThousandth &&
            !accent2SatMod.HasGreenModulationThousandth &&
            !accent2SatMod.HasGreenOffsetThousandth &&
            !accent2SatMod.HasBlueModulationThousandth &&
            !accent2SatMod.HasBlueOffsetThousandth &&
            !accent2SatMod.HasHueModulationThousandth &&
            !accent2SatMod.HasHueOffsetAngleThousandth &&
            !accent2SatMod.HasGray && !accent2SatMod.HasComp && !accent2SatMod.HasInv &&
            !accent2SatMod.HasGamma && !accent2SatMod.HasInvGamma &&
            accent2SatMod.SaturationModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2SatMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 satMod.",
                    PartPath(themePart));
            var requested = (long)accent2SatMod.SaturationModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var satMod = color.FirstChild as A.SaturationModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 satMod.",
                        PartPath(themePart));
                satMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasRedModulationThousandth: true } accent2RedMod &&
            !accent2RedMod.HasTintThousandth &&
            !accent2RedMod.HasShadeThousandth &&
            !accent2RedMod.HasLuminanceModulationThousandth &&
            !accent2RedMod.HasLuminanceOffsetThousandth &&
            !accent2RedMod.HasAlphaModulationThousandth &&
            !accent2RedMod.HasAlphaOffsetThousandth &&
            !accent2RedMod.HasSaturationModulationThousandth &&
            !accent2RedMod.HasSaturationOffsetThousandth &&
            !accent2RedMod.HasRedOffsetThousandth &&
            !accent2RedMod.HasGreenModulationThousandth &&
            !accent2RedMod.HasGreenOffsetThousandth &&
            !accent2RedMod.HasBlueModulationThousandth &&
            !accent2RedMod.HasBlueOffsetThousandth &&
            !accent2RedMod.HasHueModulationThousandth &&
            !accent2RedMod.HasHueOffsetAngleThousandth &&
            !accent2RedMod.HasGray && !accent2RedMod.HasComp && !accent2RedMod.HasInv &&
            !accent2RedMod.HasGamma && !accent2RedMod.HasInvGamma &&
            accent2RedMod.RedModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2RedMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 redMod.",
                    PartPath(themePart));
            var requested = (long)accent2RedMod.RedModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var redMod = color.FirstChild as A.RedModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 redMod.",
                        PartPath(themePart));
                redMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasSaturationOffsetThousandth: true } accent2SatOff &&
            !accent2SatOff.HasTintThousandth &&
            !accent2SatOff.HasShadeThousandth &&
            !accent2SatOff.HasLuminanceModulationThousandth &&
            !accent2SatOff.HasLuminanceOffsetThousandth &&
            !accent2SatOff.HasAlphaModulationThousandth &&
            !accent2SatOff.HasAlphaOffsetThousandth &&
            !accent2SatOff.HasSaturationModulationThousandth &&
            !accent2SatOff.HasRedModulationThousandth &&
            !accent2SatOff.HasRedOffsetThousandth &&
            !accent2SatOff.HasGreenModulationThousandth &&
            !accent2SatOff.HasGreenOffsetThousandth &&
            !accent2SatOff.HasBlueModulationThousandth &&
            !accent2SatOff.HasBlueOffsetThousandth &&
            !accent2SatOff.HasHueModulationThousandth &&
            !accent2SatOff.HasHueOffsetAngleThousandth &&
            !accent2SatOff.HasGray && !accent2SatOff.HasComp && !accent2SatOff.HasInv &&
            !accent2SatOff.HasGamma && !accent2SatOff.HasInvGamma &&
            accent2SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent2SatOff.SaturationOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2SatOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 satOff.",
                    PartPath(themePart));
            var requested = (long)accent2SatOff.SaturationOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var satOff = color.FirstChild as A.SaturationOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 satOff.",
                        PartPath(themePart));
                satOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasRedOffsetThousandth: true } accent2RedOff &&
            !accent2RedOff.HasTintThousandth &&
            !accent2RedOff.HasShadeThousandth &&
            !accent2RedOff.HasLuminanceModulationThousandth &&
            !accent2RedOff.HasLuminanceOffsetThousandth &&
            !accent2RedOff.HasAlphaModulationThousandth &&
            !accent2RedOff.HasAlphaOffsetThousandth &&
            !accent2RedOff.HasSaturationModulationThousandth &&
            !accent2RedOff.HasSaturationOffsetThousandth &&
            !accent2RedOff.HasRedModulationThousandth &&
            !accent2RedOff.HasGreenModulationThousandth &&
            !accent2RedOff.HasGreenOffsetThousandth &&
            !accent2RedOff.HasBlueModulationThousandth &&
            !accent2RedOff.HasBlueOffsetThousandth &&
            !accent2RedOff.HasHueModulationThousandth &&
            !accent2RedOff.HasHueOffsetAngleThousandth &&
            !accent2RedOff.HasGray && !accent2RedOff.HasComp && !accent2RedOff.HasInv &&
            !accent2RedOff.HasGamma && !accent2RedOff.HasInvGamma &&
            accent2RedOff.RedOffsetThousandth >= -100_000 &&
            accent2RedOff.RedOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2RedOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 redOff.",
                    PartPath(themePart));
            var requested = (long)accent2RedOff.RedOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var redOff = color.FirstChild as A.RedOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 redOff.",
                        PartPath(themePart));
                redOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasGreenModulationThousandth: true } accent2GreenMod &&
            !accent2GreenMod.HasTintThousandth &&
            !accent2GreenMod.HasShadeThousandth &&
            !accent2GreenMod.HasLuminanceModulationThousandth &&
            !accent2GreenMod.HasLuminanceOffsetThousandth &&
            !accent2GreenMod.HasAlphaModulationThousandth &&
            !accent2GreenMod.HasAlphaOffsetThousandth &&
            !accent2GreenMod.HasSaturationModulationThousandth &&
            !accent2GreenMod.HasSaturationOffsetThousandth &&
            !accent2GreenMod.HasRedModulationThousandth &&
            !accent2GreenMod.HasRedOffsetThousandth &&
            !accent2GreenMod.HasGreenOffsetThousandth &&
            !accent2GreenMod.HasBlueModulationThousandth &&
            !accent2GreenMod.HasBlueOffsetThousandth &&
            !accent2GreenMod.HasHueModulationThousandth &&
            !accent2GreenMod.HasHueOffsetAngleThousandth &&
            !accent2GreenMod.HasGray && !accent2GreenMod.HasComp && !accent2GreenMod.HasInv &&
            !accent2GreenMod.HasGamma && !accent2GreenMod.HasInvGamma &&
            accent2GreenMod.GreenModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2GreenMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 greenMod.",
                    PartPath(themePart));
            var requested = (long)accent2GreenMod.GreenModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var greenMod = color.FirstChild as A.GreenModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 greenMod.",
                        PartPath(themePart));
                greenMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasGreenOffsetThousandth: true } accent2GreenOff &&
            !accent2GreenOff.HasTintThousandth &&
            !accent2GreenOff.HasShadeThousandth &&
            !accent2GreenOff.HasLuminanceModulationThousandth &&
            !accent2GreenOff.HasLuminanceOffsetThousandth &&
            !accent2GreenOff.HasAlphaModulationThousandth &&
            !accent2GreenOff.HasAlphaOffsetThousandth &&
            !accent2GreenOff.HasSaturationModulationThousandth &&
            !accent2GreenOff.HasSaturationOffsetThousandth &&
            !accent2GreenOff.HasRedModulationThousandth &&
            !accent2GreenOff.HasRedOffsetThousandth &&
            !accent2GreenOff.HasGreenModulationThousandth &&
            !accent2GreenOff.HasBlueModulationThousandth &&
            !accent2GreenOff.HasBlueOffsetThousandth &&
            !accent2GreenOff.HasHueModulationThousandth &&
            !accent2GreenOff.HasHueOffsetAngleThousandth &&
            !accent2GreenOff.HasGray && !accent2GreenOff.HasComp && !accent2GreenOff.HasInv &&
            !accent2GreenOff.HasGamma && !accent2GreenOff.HasInvGamma &&
            accent2GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent2GreenOff.GreenOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2GreenOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 greenOff.",
                    PartPath(themePart));
            var requested = (long)accent2GreenOff.GreenOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var greenOff = color.FirstChild as A.GreenOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 greenOff.",
                        PartPath(themePart));
                greenOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasBlueModulationThousandth: true } accent2BlueMod &&
            !accent2BlueMod.HasTintThousandth &&
            !accent2BlueMod.HasShadeThousandth &&
            !accent2BlueMod.HasLuminanceModulationThousandth &&
            !accent2BlueMod.HasLuminanceOffsetThousandth &&
            !accent2BlueMod.HasAlphaModulationThousandth &&
            !accent2BlueMod.HasAlphaOffsetThousandth &&
            !accent2BlueMod.HasSaturationModulationThousandth &&
            !accent2BlueMod.HasSaturationOffsetThousandth &&
            !accent2BlueMod.HasRedModulationThousandth &&
            !accent2BlueMod.HasRedOffsetThousandth &&
            !accent2BlueMod.HasGreenModulationThousandth &&
            !accent2BlueMod.HasGreenOffsetThousandth &&
            !accent2BlueMod.HasBlueOffsetThousandth &&
            !accent2BlueMod.HasHueModulationThousandth &&
            !accent2BlueMod.HasHueOffsetAngleThousandth &&
            !accent2BlueMod.HasGray && !accent2BlueMod.HasComp && !accent2BlueMod.HasInv &&
            !accent2BlueMod.HasGamma && !accent2BlueMod.HasInvGamma &&
            accent2BlueMod.BlueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2BlueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 blueMod.",
                    PartPath(themePart));
            var requested = (long)accent2BlueMod.BlueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var blueMod = color.FirstChild as A.BlueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 blueMod.",
                        PartPath(themePart));
                blueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasBlueOffsetThousandth: true } accent2BlueOff &&
            !accent2BlueOff.HasTintThousandth &&
            !accent2BlueOff.HasShadeThousandth &&
            !accent2BlueOff.HasLuminanceModulationThousandth &&
            !accent2BlueOff.HasLuminanceOffsetThousandth &&
            !accent2BlueOff.HasAlphaModulationThousandth &&
            !accent2BlueOff.HasAlphaOffsetThousandth &&
            !accent2BlueOff.HasSaturationModulationThousandth &&
            !accent2BlueOff.HasSaturationOffsetThousandth &&
            !accent2BlueOff.HasRedModulationThousandth &&
            !accent2BlueOff.HasRedOffsetThousandth &&
            !accent2BlueOff.HasGreenModulationThousandth &&
            !accent2BlueOff.HasGreenOffsetThousandth &&
            !accent2BlueOff.HasBlueModulationThousandth &&
            !accent2BlueOff.HasHueModulationThousandth &&
            !accent2BlueOff.HasHueOffsetAngleThousandth &&
            !accent2BlueOff.HasGray && !accent2BlueOff.HasComp && !accent2BlueOff.HasInv &&
            !accent2BlueOff.HasGamma && !accent2BlueOff.HasInvGamma &&
            accent2BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent2BlueOff.BlueOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2BlueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 blueOff.",
                    PartPath(themePart));
            var requested = (long)accent2BlueOff.BlueOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var blueOff = color.FirstChild as A.BlueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 blueOff.",
                        PartPath(themePart));
                blueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasHueModulationThousandth: true } accent2HueMod &&
            !accent2HueMod.HasTintThousandth &&
            !accent2HueMod.HasShadeThousandth &&
            !accent2HueMod.HasLuminanceModulationThousandth &&
            !accent2HueMod.HasLuminanceOffsetThousandth &&
            !accent2HueMod.HasAlphaModulationThousandth &&
            !accent2HueMod.HasAlphaOffsetThousandth &&
            !accent2HueMod.HasSaturationModulationThousandth &&
            !accent2HueMod.HasSaturationOffsetThousandth &&
            !accent2HueMod.HasRedModulationThousandth &&
            !accent2HueMod.HasRedOffsetThousandth &&
            !accent2HueMod.HasGreenModulationThousandth &&
            !accent2HueMod.HasGreenOffsetThousandth &&
            !accent2HueMod.HasBlueModulationThousandth &&
            !accent2HueMod.HasHueOffsetAngleThousandth &&
            !accent2HueMod.HasGray && !accent2HueMod.HasComp && !accent2HueMod.HasInv &&
            !accent2HueMod.HasGamma && !accent2HueMod.HasInvGamma &&
            accent2HueMod.HueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2HueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 hueMod.",
                    PartPath(themePart));
            var requested = (long)accent2HueMod.HueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var hueMod = color.FirstChild as A.HueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 hueMod.",
                        PartPath(themePart));
                hueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent2", HasHueOffsetAngleThousandth: true } accent2HueOff &&
            !accent2HueOff.HasTintThousandth &&
            !accent2HueOff.HasShadeThousandth &&
            !accent2HueOff.HasLuminanceModulationThousandth &&
            !accent2HueOff.HasLuminanceOffsetThousandth &&
            !accent2HueOff.HasAlphaModulationThousandth &&
            !accent2HueOff.HasAlphaOffsetThousandth &&
            !accent2HueOff.HasSaturationModulationThousandth &&
            !accent2HueOff.HasSaturationOffsetThousandth &&
            !accent2HueOff.HasRedModulationThousandth &&
            !accent2HueOff.HasRedOffsetThousandth &&
            !accent2HueOff.HasGreenModulationThousandth &&
            !accent2HueOff.HasGreenOffsetThousandth &&
            !accent2HueOff.HasBlueModulationThousandth &&
            !accent2HueOff.HasHueModulationThousandth &&
            !accent2HueOff.HasGray && !accent2HueOff.HasComp && !accent2HueOff.HasInv &&
            !accent2HueOff.HasGamma && !accent2HueOff.HasInvGamma &&
            accent2HueOff.HueOffsetAngleThousandth >= -21_600_000 &&
            accent2HueOff.HueOffsetAngleThousandth <= 21_600_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent2HueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent2 hueOff.",
                    PartPath(themePart));
            var requested = (long)accent2HueOff.HueOffsetAngleThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent2Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 color.",
                        PartPath(themePart));
                var hueOff = color.FirstChild as A.HueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent2 hueOff.",
                        PartPath(themePart));
                hueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasTintThousandth: true } accent3Tint &&
            !accent3Tint.HasShadeThousandth &&
            !accent3Tint.HasLuminanceModulationThousandth &&
            !accent3Tint.HasLuminanceOffsetThousandth &&
            !accent3Tint.HasAlphaModulationThousandth &&
            !accent3Tint.HasAlphaOffsetThousandth &&
            !accent3Tint.HasSaturationModulationThousandth &&
            !accent3Tint.HasSaturationOffsetThousandth &&
            !accent3Tint.HasRedModulationThousandth &&
            !accent3Tint.HasRedOffsetThousandth &&
            !accent3Tint.HasGreenModulationThousandth &&
            !accent3Tint.HasGreenOffsetThousandth &&
            !accent3Tint.HasBlueModulationThousandth &&
            !accent3Tint.HasBlueOffsetThousandth &&
            !accent3Tint.HasHueModulationThousandth &&
            !accent3Tint.HasHueOffsetAngleThousandth &&
            !accent3Tint.HasGray && !accent3Tint.HasComp && !accent3Tint.HasInv &&
            !accent3Tint.HasGamma && !accent3Tint.HasInvGamma &&
            accent3Tint.TintThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3Tint(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 tint.",
                    PartPath(themePart));
            var requested = (long)accent3Tint.TintThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var tint = color.FirstChild as A.Tint ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 tint.",
                        PartPath(themePart));
                tint.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasShadeThousandth: true } accent3Shade &&
            !accent3Shade.HasTintThousandth &&
            !accent3Shade.HasLuminanceModulationThousandth &&
            !accent3Shade.HasLuminanceOffsetThousandth &&
            !accent3Shade.HasAlphaModulationThousandth &&
            !accent3Shade.HasAlphaOffsetThousandth &&
            !accent3Shade.HasSaturationModulationThousandth &&
            !accent3Shade.HasSaturationOffsetThousandth &&
            !accent3Shade.HasRedModulationThousandth &&
            !accent3Shade.HasRedOffsetThousandth &&
            !accent3Shade.HasGreenModulationThousandth &&
            !accent3Shade.HasGreenOffsetThousandth &&
            !accent3Shade.HasBlueModulationThousandth &&
            !accent3Shade.HasBlueOffsetThousandth &&
            !accent3Shade.HasHueModulationThousandth &&
            !accent3Shade.HasHueOffsetAngleThousandth &&
            !accent3Shade.HasGray && !accent3Shade.HasComp && !accent3Shade.HasInv &&
            !accent3Shade.HasGamma && !accent3Shade.HasInvGamma &&
            accent3Shade.ShadeThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3Shade(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 shade.",
                    PartPath(themePart));
            var requested = (long)accent3Shade.ShadeThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var shade = color.FirstChild as A.Shade ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 shade.",
                        PartPath(themePart));
                shade.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasLuminanceModulationThousandth: true } accent3LumMod &&
            !accent3LumMod.HasTintThousandth &&
            !accent3LumMod.HasShadeThousandth &&
            !accent3LumMod.HasLuminanceOffsetThousandth &&
            !accent3LumMod.HasAlphaModulationThousandth &&
            !accent3LumMod.HasAlphaOffsetThousandth &&
            !accent3LumMod.HasSaturationModulationThousandth &&
            !accent3LumMod.HasSaturationOffsetThousandth &&
            !accent3LumMod.HasRedModulationThousandth &&
            !accent3LumMod.HasRedOffsetThousandth &&
            !accent3LumMod.HasGreenModulationThousandth &&
            !accent3LumMod.HasGreenOffsetThousandth &&
            !accent3LumMod.HasBlueModulationThousandth &&
            !accent3LumMod.HasBlueOffsetThousandth &&
            !accent3LumMod.HasHueModulationThousandth &&
            !accent3LumMod.HasHueOffsetAngleThousandth &&
            !accent3LumMod.HasGray && !accent3LumMod.HasComp && !accent3LumMod.HasInv &&
            !accent3LumMod.HasGamma && !accent3LumMod.HasInvGamma &&
            accent3LumMod.LuminanceModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3LumMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 lumMod.",
                    PartPath(themePart));
            var requested = (long)accent3LumMod.LuminanceModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var lumMod = color.FirstChild as A.LuminanceModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 lumMod.",
                        PartPath(themePart));
                lumMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasLuminanceOffsetThousandth: true } accent3LumOff &&
            !accent3LumOff.HasTintThousandth &&
            !accent3LumOff.HasShadeThousandth &&
            !accent3LumOff.HasLuminanceModulationThousandth &&
            !accent3LumOff.HasAlphaModulationThousandth &&
            !accent3LumOff.HasAlphaOffsetThousandth &&
            !accent3LumOff.HasSaturationModulationThousandth &&
            !accent3LumOff.HasSaturationOffsetThousandth &&
            !accent3LumOff.HasRedModulationThousandth &&
            !accent3LumOff.HasRedOffsetThousandth &&
            !accent3LumOff.HasGreenModulationThousandth &&
            !accent3LumOff.HasGreenOffsetThousandth &&
            !accent3LumOff.HasBlueModulationThousandth &&
            !accent3LumOff.HasBlueOffsetThousandth &&
            !accent3LumOff.HasHueModulationThousandth &&
            !accent3LumOff.HasHueOffsetAngleThousandth &&
            !accent3LumOff.HasGray && !accent3LumOff.HasComp && !accent3LumOff.HasInv &&
            !accent3LumOff.HasGamma && !accent3LumOff.HasInvGamma &&
            accent3LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent3LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3LumOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 lumOff.",
                    PartPath(themePart));
            var requested = (long)accent3LumOff.LuminanceOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var lumOff = color.FirstChild as A.LuminanceOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 lumOff.",
                        PartPath(themePart));
                lumOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasAlphaModulationThousandth: true } accent3AlphaMod &&
            !accent3AlphaMod.HasTintThousandth &&
            !accent3AlphaMod.HasShadeThousandth &&
            !accent3AlphaMod.HasLuminanceModulationThousandth &&
            !accent3AlphaMod.HasLuminanceOffsetThousandth &&
            !accent3AlphaMod.HasAlphaOffsetThousandth &&
            !accent3AlphaMod.HasSaturationModulationThousandth &&
            !accent3AlphaMod.HasSaturationOffsetThousandth &&
            !accent3AlphaMod.HasRedModulationThousandth &&
            !accent3AlphaMod.HasRedOffsetThousandth &&
            !accent3AlphaMod.HasGreenModulationThousandth &&
            !accent3AlphaMod.HasGreenOffsetThousandth &&
            !accent3AlphaMod.HasBlueModulationThousandth &&
            !accent3AlphaMod.HasBlueOffsetThousandth &&
            !accent3AlphaMod.HasHueModulationThousandth &&
            !accent3AlphaMod.HasHueOffsetAngleThousandth &&
            !accent3AlphaMod.HasGray && !accent3AlphaMod.HasComp && !accent3AlphaMod.HasInv &&
            !accent3AlphaMod.HasGamma && !accent3AlphaMod.HasInvGamma &&
            accent3AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3AlphaMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 alphaMod.",
                    PartPath(themePart));
            var requested = (long)accent3AlphaMod.AlphaModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var alphaMod = color.FirstChild as A.AlphaModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 alphaMod.",
                        PartPath(themePart));
                alphaMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasAlphaOffsetThousandth: true } accent3AlphaOff &&
            !accent3AlphaOff.HasTintThousandth &&
            !accent3AlphaOff.HasShadeThousandth &&
            !accent3AlphaOff.HasLuminanceModulationThousandth &&
            !accent3AlphaOff.HasLuminanceOffsetThousandth &&
            !accent3AlphaOff.HasAlphaModulationThousandth &&
            !accent3AlphaOff.HasSaturationModulationThousandth &&
            !accent3AlphaOff.HasSaturationOffsetThousandth &&
            !accent3AlphaOff.HasRedModulationThousandth &&
            !accent3AlphaOff.HasRedOffsetThousandth &&
            !accent3AlphaOff.HasGreenModulationThousandth &&
            !accent3AlphaOff.HasGreenOffsetThousandth &&
            !accent3AlphaOff.HasBlueModulationThousandth &&
            !accent3AlphaOff.HasBlueOffsetThousandth &&
            !accent3AlphaOff.HasHueModulationThousandth &&
            !accent3AlphaOff.HasHueOffsetAngleThousandth &&
            !accent3AlphaOff.HasGray && !accent3AlphaOff.HasComp && !accent3AlphaOff.HasInv &&
            !accent3AlphaOff.HasGamma && !accent3AlphaOff.HasInvGamma &&
            accent3AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent3AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3AlphaOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 alphaOff.",
                    PartPath(themePart));
            var requested = (long)accent3AlphaOff.AlphaOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var alphaOff = color.FirstChild as A.AlphaOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 alphaOff.",
                        PartPath(themePart));
                alphaOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasSaturationModulationThousandth: true } accent3SatMod &&
            !accent3SatMod.HasTintThousandth &&
            !accent3SatMod.HasShadeThousandth &&
            !accent3SatMod.HasLuminanceModulationThousandth &&
            !accent3SatMod.HasLuminanceOffsetThousandth &&
            !accent3SatMod.HasAlphaModulationThousandth &&
            !accent3SatMod.HasAlphaOffsetThousandth &&
            !accent3SatMod.HasSaturationOffsetThousandth &&
            !accent3SatMod.HasRedModulationThousandth &&
            !accent3SatMod.HasRedOffsetThousandth &&
            !accent3SatMod.HasGreenModulationThousandth &&
            !accent3SatMod.HasGreenOffsetThousandth &&
            !accent3SatMod.HasBlueModulationThousandth &&
            !accent3SatMod.HasBlueOffsetThousandth &&
            !accent3SatMod.HasHueModulationThousandth &&
            !accent3SatMod.HasHueOffsetAngleThousandth &&
            !accent3SatMod.HasGray && !accent3SatMod.HasComp && !accent3SatMod.HasInv &&
            !accent3SatMod.HasGamma && !accent3SatMod.HasInvGamma &&
            accent3SatMod.SaturationModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3SatMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 satMod.",
                    PartPath(themePart));
            var requested = (long)accent3SatMod.SaturationModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var satMod = color.FirstChild as A.SaturationModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 satMod.",
                        PartPath(themePart));
                satMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasSaturationOffsetThousandth: true } accent3SatOff &&
            !accent3SatOff.HasTintThousandth &&
            !accent3SatOff.HasShadeThousandth &&
            !accent3SatOff.HasLuminanceModulationThousandth &&
            !accent3SatOff.HasLuminanceOffsetThousandth &&
            !accent3SatOff.HasAlphaModulationThousandth &&
            !accent3SatOff.HasAlphaOffsetThousandth &&
            !accent3SatOff.HasSaturationModulationThousandth &&
            !accent3SatOff.HasRedModulationThousandth &&
            !accent3SatOff.HasRedOffsetThousandth &&
            !accent3SatOff.HasGreenModulationThousandth &&
            !accent3SatOff.HasGreenOffsetThousandth &&
            !accent3SatOff.HasBlueModulationThousandth &&
            !accent3SatOff.HasBlueOffsetThousandth &&
            !accent3SatOff.HasHueModulationThousandth &&
            !accent3SatOff.HasHueOffsetAngleThousandth &&
            !accent3SatOff.HasGray && !accent3SatOff.HasComp && !accent3SatOff.HasInv &&
            !accent3SatOff.HasGamma && !accent3SatOff.HasInvGamma &&
            accent3SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent3SatOff.SaturationOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3SatOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 satOff.",
                    PartPath(themePart));
            var requested = (long)accent3SatOff.SaturationOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var satOff = color.FirstChild as A.SaturationOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 satOff.",
                        PartPath(themePart));
                satOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasRedModulationThousandth: true } accent3RedMod &&
            !accent3RedMod.HasTintThousandth &&
            !accent3RedMod.HasShadeThousandth &&
            !accent3RedMod.HasLuminanceModulationThousandth &&
            !accent3RedMod.HasLuminanceOffsetThousandth &&
            !accent3RedMod.HasAlphaModulationThousandth &&
            !accent3RedMod.HasAlphaOffsetThousandth &&
            !accent3RedMod.HasSaturationModulationThousandth &&
            !accent3RedMod.HasSaturationOffsetThousandth &&
            !accent3RedMod.HasRedOffsetThousandth &&
            !accent3RedMod.HasGreenModulationThousandth &&
            !accent3RedMod.HasGreenOffsetThousandth &&
            !accent3RedMod.HasBlueModulationThousandth &&
            !accent3RedMod.HasBlueOffsetThousandth &&
            !accent3RedMod.HasHueModulationThousandth &&
            !accent3RedMod.HasHueOffsetAngleThousandth &&
            !accent3RedMod.HasGray && !accent3RedMod.HasComp && !accent3RedMod.HasInv &&
            !accent3RedMod.HasGamma && !accent3RedMod.HasInvGamma &&
            accent3RedMod.RedModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3RedMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 redMod.",
                    PartPath(themePart));
            var requested = (long)accent3RedMod.RedModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var redMod = color.FirstChild as A.RedModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 redMod.",
                        PartPath(themePart));
                redMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasRedOffsetThousandth: true } accent3RedOff &&
            !accent3RedOff.HasTintThousandth &&
            !accent3RedOff.HasShadeThousandth &&
            !accent3RedOff.HasLuminanceModulationThousandth &&
            !accent3RedOff.HasLuminanceOffsetThousandth &&
            !accent3RedOff.HasAlphaModulationThousandth &&
            !accent3RedOff.HasAlphaOffsetThousandth &&
            !accent3RedOff.HasSaturationModulationThousandth &&
            !accent3RedOff.HasSaturationOffsetThousandth &&
            !accent3RedOff.HasRedModulationThousandth &&
            !accent3RedOff.HasGreenModulationThousandth &&
            !accent3RedOff.HasGreenOffsetThousandth &&
            !accent3RedOff.HasBlueModulationThousandth &&
            !accent3RedOff.HasBlueOffsetThousandth &&
            !accent3RedOff.HasHueModulationThousandth &&
            !accent3RedOff.HasHueOffsetAngleThousandth &&
            !accent3RedOff.HasGray && !accent3RedOff.HasComp && !accent3RedOff.HasInv &&
            !accent3RedOff.HasGamma && !accent3RedOff.HasInvGamma &&
            accent3RedOff.RedOffsetThousandth >= -100_000 &&
            accent3RedOff.RedOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3RedOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 redOff.",
                    PartPath(themePart));
            var requested = (long)accent3RedOff.RedOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var redOff = color.FirstChild as A.RedOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 redOff.",
                        PartPath(themePart));
                redOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasGreenModulationThousandth: true } accent3GreenMod &&
            !accent3GreenMod.HasTintThousandth &&
            !accent3GreenMod.HasShadeThousandth &&
            !accent3GreenMod.HasLuminanceModulationThousandth &&
            !accent3GreenMod.HasLuminanceOffsetThousandth &&
            !accent3GreenMod.HasAlphaModulationThousandth &&
            !accent3GreenMod.HasAlphaOffsetThousandth &&
            !accent3GreenMod.HasSaturationModulationThousandth &&
            !accent3GreenMod.HasSaturationOffsetThousandth &&
            !accent3GreenMod.HasRedModulationThousandth &&
            !accent3GreenMod.HasRedOffsetThousandth &&
            !accent3GreenMod.HasGreenOffsetThousandth &&
            !accent3GreenMod.HasBlueModulationThousandth &&
            !accent3GreenMod.HasBlueOffsetThousandth &&
            !accent3GreenMod.HasHueModulationThousandth &&
            !accent3GreenMod.HasHueOffsetAngleThousandth &&
            !accent3GreenMod.HasGray && !accent3GreenMod.HasComp && !accent3GreenMod.HasInv &&
            !accent3GreenMod.HasGamma && !accent3GreenMod.HasInvGamma &&
            accent3GreenMod.GreenModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3GreenMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 greenMod.",
                    PartPath(themePart));
            var requested = (long)accent3GreenMod.GreenModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var greenMod = color.FirstChild as A.GreenModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 greenMod.",
                        PartPath(themePart));
                greenMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasGreenOffsetThousandth: true } accent3GreenOff &&
            !accent3GreenOff.HasTintThousandth &&
            !accent3GreenOff.HasShadeThousandth &&
            !accent3GreenOff.HasLuminanceModulationThousandth &&
            !accent3GreenOff.HasLuminanceOffsetThousandth &&
            !accent3GreenOff.HasAlphaModulationThousandth &&
            !accent3GreenOff.HasAlphaOffsetThousandth &&
            !accent3GreenOff.HasSaturationModulationThousandth &&
            !accent3GreenOff.HasSaturationOffsetThousandth &&
            !accent3GreenOff.HasRedModulationThousandth &&
            !accent3GreenOff.HasRedOffsetThousandth &&
            !accent3GreenOff.HasGreenModulationThousandth &&
            !accent3GreenOff.HasBlueModulationThousandth &&
            !accent3GreenOff.HasBlueOffsetThousandth &&
            !accent3GreenOff.HasHueModulationThousandth &&
            !accent3GreenOff.HasHueOffsetAngleThousandth &&
            !accent3GreenOff.HasGray && !accent3GreenOff.HasComp && !accent3GreenOff.HasInv &&
            !accent3GreenOff.HasGamma && !accent3GreenOff.HasInvGamma &&
            accent3GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent3GreenOff.GreenOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3GreenOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 greenOff.",
                    PartPath(themePart));
            var requested = (long)accent3GreenOff.GreenOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var greenOff = color.FirstChild as A.GreenOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 greenOff.",
                        PartPath(themePart));
                greenOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasBlueModulationThousandth: true } accent3BlueMod &&
            !accent3BlueMod.HasTintThousandth &&
            !accent3BlueMod.HasShadeThousandth &&
            !accent3BlueMod.HasLuminanceModulationThousandth &&
            !accent3BlueMod.HasLuminanceOffsetThousandth &&
            !accent3BlueMod.HasAlphaModulationThousandth &&
            !accent3BlueMod.HasAlphaOffsetThousandth &&
            !accent3BlueMod.HasSaturationModulationThousandth &&
            !accent3BlueMod.HasSaturationOffsetThousandth &&
            !accent3BlueMod.HasRedModulationThousandth &&
            !accent3BlueMod.HasRedOffsetThousandth &&
            !accent3BlueMod.HasGreenModulationThousandth &&
            !accent3BlueMod.HasGreenOffsetThousandth &&
            !accent3BlueMod.HasBlueOffsetThousandth &&
            !accent3BlueMod.HasHueModulationThousandth &&
            !accent3BlueMod.HasHueOffsetAngleThousandth &&
            !accent3BlueMod.HasGray && !accent3BlueMod.HasComp && !accent3BlueMod.HasInv &&
            !accent3BlueMod.HasGamma && !accent3BlueMod.HasInvGamma &&
            accent3BlueMod.BlueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3BlueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 blueMod.",
                    PartPath(themePart));
            var requested = (long)accent3BlueMod.BlueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var blueMod = color.FirstChild as A.BlueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 blueMod.",
                        PartPath(themePart));
                blueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasBlueOffsetThousandth: true } accent3BlueOff &&
            !accent3BlueOff.HasTintThousandth &&
            !accent3BlueOff.HasShadeThousandth &&
            !accent3BlueOff.HasLuminanceModulationThousandth &&
            !accent3BlueOff.HasLuminanceOffsetThousandth &&
            !accent3BlueOff.HasAlphaModulationThousandth &&
            !accent3BlueOff.HasAlphaOffsetThousandth &&
            !accent3BlueOff.HasSaturationModulationThousandth &&
            !accent3BlueOff.HasSaturationOffsetThousandth &&
            !accent3BlueOff.HasRedModulationThousandth &&
            !accent3BlueOff.HasRedOffsetThousandth &&
            !accent3BlueOff.HasGreenModulationThousandth &&
            !accent3BlueOff.HasGreenOffsetThousandth &&
            !accent3BlueOff.HasBlueModulationThousandth &&
            !accent3BlueOff.HasHueModulationThousandth &&
            !accent3BlueOff.HasHueOffsetAngleThousandth &&
            !accent3BlueOff.HasGray && !accent3BlueOff.HasComp && !accent3BlueOff.HasInv &&
            !accent3BlueOff.HasGamma && !accent3BlueOff.HasInvGamma &&
            accent3BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent3BlueOff.BlueOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3BlueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 blueOff.",
                    PartPath(themePart));
            var requested = (long)accent3BlueOff.BlueOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var blueOff = color.FirstChild as A.BlueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 blueOff.",
                        PartPath(themePart));
                blueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasHueModulationThousandth: true } accent3HueMod &&
            !accent3HueMod.HasTintThousandth &&
            !accent3HueMod.HasShadeThousandth &&
            !accent3HueMod.HasLuminanceModulationThousandth &&
            !accent3HueMod.HasLuminanceOffsetThousandth &&
            !accent3HueMod.HasAlphaModulationThousandth &&
            !accent3HueMod.HasAlphaOffsetThousandth &&
            !accent3HueMod.HasSaturationModulationThousandth &&
            !accent3HueMod.HasSaturationOffsetThousandth &&
            !accent3HueMod.HasRedModulationThousandth &&
            !accent3HueMod.HasRedOffsetThousandth &&
            !accent3HueMod.HasGreenModulationThousandth &&
            !accent3HueMod.HasGreenOffsetThousandth &&
            !accent3HueMod.HasBlueModulationThousandth &&
            !accent3HueMod.HasHueOffsetAngleThousandth &&
            !accent3HueMod.HasGray && !accent3HueMod.HasComp && !accent3HueMod.HasInv &&
            !accent3HueMod.HasGamma && !accent3HueMod.HasInvGamma &&
            accent3HueMod.HueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3HueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 hueMod.",
                    PartPath(themePart));
            var requested = (long)accent3HueMod.HueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var hueMod = color.FirstChild as A.HueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 hueMod.",
                        PartPath(themePart));
                hueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent3", HasHueOffsetAngleThousandth: true } accent3HueOff &&
            !accent3HueOff.HasTintThousandth &&
            !accent3HueOff.HasShadeThousandth &&
            !accent3HueOff.HasLuminanceModulationThousandth &&
            !accent3HueOff.HasLuminanceOffsetThousandth &&
            !accent3HueOff.HasAlphaModulationThousandth &&
            !accent3HueOff.HasAlphaOffsetThousandth &&
            !accent3HueOff.HasSaturationModulationThousandth &&
            !accent3HueOff.HasSaturationOffsetThousandth &&
            !accent3HueOff.HasRedModulationThousandth &&
            !accent3HueOff.HasRedOffsetThousandth &&
            !accent3HueOff.HasGreenModulationThousandth &&
            !accent3HueOff.HasGreenOffsetThousandth &&
            !accent3HueOff.HasBlueModulationThousandth &&
            !accent3HueOff.HasHueModulationThousandth &&
            !accent3HueOff.HasGray && !accent3HueOff.HasComp && !accent3HueOff.HasInv &&
            !accent3HueOff.HasGamma && !accent3HueOff.HasInvGamma &&
            accent3HueOff.HueOffsetAngleThousandth >= -21_600_000 &&
            accent3HueOff.HueOffsetAngleThousandth <= 21_600_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent3HueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent3 hueOff.",
                    PartPath(themePart));
            var requested = (long)accent3HueOff.HueOffsetAngleThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent3Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 color.",
                        PartPath(themePart));
                var hueOff = color.FirstChild as A.HueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent3 hueOff.",
                        PartPath(themePart));
                hueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasHueOffsetAngleThousandth: true } accent4HueOff &&
            !accent4HueOff.HasTintThousandth &&
            !accent4HueOff.HasShadeThousandth &&
            !accent4HueOff.HasLuminanceModulationThousandth &&
            !accent4HueOff.HasLuminanceOffsetThousandth &&
            !accent4HueOff.HasAlphaModulationThousandth &&
            !accent4HueOff.HasAlphaOffsetThousandth &&
            !accent4HueOff.HasSaturationModulationThousandth &&
            !accent4HueOff.HasSaturationOffsetThousandth &&
            !accent4HueOff.HasRedModulationThousandth &&
            !accent4HueOff.HasRedOffsetThousandth &&
            !accent4HueOff.HasGreenModulationThousandth &&
            !accent4HueOff.HasGreenOffsetThousandth &&
            !accent4HueOff.HasBlueModulationThousandth &&
            !accent4HueOff.HasHueModulationThousandth &&
            !accent4HueOff.HasGray && !accent4HueOff.HasComp && !accent4HueOff.HasInv &&
            !accent4HueOff.HasGamma && !accent4HueOff.HasInvGamma &&
            accent4HueOff.HueOffsetAngleThousandth >= -21_600_000 &&
            accent4HueOff.HueOffsetAngleThousandth <= 21_600_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4HueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 hueOff.",
                    PartPath(themePart));
            var requested = (long)accent4HueOff.HueOffsetAngleThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var hueOff = color.FirstChild as A.HueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 hueOff.",
                        PartPath(themePart));
                hueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasTintThousandth: true } accent5Tint &&
            !accent5Tint.HasShadeThousandth &&
            !accent5Tint.HasLuminanceModulationThousandth &&
            !accent5Tint.HasLuminanceOffsetThousandth &&
            !accent5Tint.HasAlphaModulationThousandth &&
            !accent5Tint.HasAlphaOffsetThousandth &&
            !accent5Tint.HasSaturationModulationThousandth &&
            !accent5Tint.HasSaturationOffsetThousandth &&
            !accent5Tint.HasRedModulationThousandth &&
            !accent5Tint.HasRedOffsetThousandth &&
            !accent5Tint.HasGreenModulationThousandth &&
            !accent5Tint.HasGreenOffsetThousandth &&
            !accent5Tint.HasBlueModulationThousandth &&
            !accent5Tint.HasBlueOffsetThousandth &&
            !accent5Tint.HasHueModulationThousandth &&
            !accent5Tint.HasHueOffsetAngleThousandth &&
            !accent5Tint.HasGray && !accent5Tint.HasComp && !accent5Tint.HasInv &&
            !accent5Tint.HasGamma && !accent5Tint.HasInvGamma &&
            accent5Tint.TintThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5Tint(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 tint.",
                    PartPath(themePart));
            var requested = (long)accent5Tint.TintThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var tint = color.FirstChild as A.Tint ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 tint.",
                        PartPath(themePart));
                tint.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent6", HasTintThousandth: true } accent6Tint &&
            !accent6Tint.HasShadeThousandth &&
            !accent6Tint.HasLuminanceModulationThousandth &&
            !accent6Tint.HasLuminanceOffsetThousandth &&
            !accent6Tint.HasAlphaModulationThousandth &&
            !accent6Tint.HasAlphaOffsetThousandth &&
            !accent6Tint.HasSaturationModulationThousandth &&
            !accent6Tint.HasSaturationOffsetThousandth &&
            !accent6Tint.HasRedModulationThousandth &&
            !accent6Tint.HasRedOffsetThousandth &&
            !accent6Tint.HasGreenModulationThousandth &&
            !accent6Tint.HasGreenOffsetThousandth &&
            !accent6Tint.HasBlueModulationThousandth &&
            !accent6Tint.HasBlueOffsetThousandth &&
            !accent6Tint.HasHueModulationThousandth &&
            !accent6Tint.HasHueOffsetAngleThousandth &&
            !accent6Tint.HasGray && !accent6Tint.HasComp && !accent6Tint.HasInv &&
            !accent6Tint.HasGamma && !accent6Tint.HasInvGamma &&
            accent6Tint.TintThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent6Tint(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent6 tint.",
                    PartPath(themePart));
            var requested = (long)accent6Tint.TintThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent6Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 color.",
                        PartPath(themePart));
                var tint = color.FirstChild as A.Tint ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 tint.",
                        PartPath(themePart));
                tint.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent6", HasShadeThousandth: true } accent6Shade &&
            !accent6Shade.HasTintThousandth &&
            !accent6Shade.HasLuminanceModulationThousandth &&
            !accent6Shade.HasLuminanceOffsetThousandth &&
            !accent6Shade.HasAlphaModulationThousandth &&
            !accent6Shade.HasAlphaOffsetThousandth &&
            !accent6Shade.HasSaturationModulationThousandth &&
            !accent6Shade.HasSaturationOffsetThousandth &&
            !accent6Shade.HasRedModulationThousandth &&
            !accent6Shade.HasRedOffsetThousandth &&
            !accent6Shade.HasGreenModulationThousandth &&
            !accent6Shade.HasGreenOffsetThousandth &&
            !accent6Shade.HasBlueModulationThousandth &&
            !accent6Shade.HasBlueOffsetThousandth &&
            !accent6Shade.HasHueModulationThousandth &&
            !accent6Shade.HasHueOffsetAngleThousandth &&
            !accent6Shade.HasGray && !accent6Shade.HasComp && !accent6Shade.HasInv &&
            !accent6Shade.HasGamma && !accent6Shade.HasInvGamma &&
            accent6Shade.ShadeThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent6Shade(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent6 shade.",
                    PartPath(themePart));
            var requested = (long)accent6Shade.ShadeThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent6Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 color.",
                        PartPath(themePart));
                var shade = color.FirstChild as A.Shade ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 shade.",
                        PartPath(themePart));
                shade.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasTintThousandth: true } accent4Tint &&
            !accent4Tint.HasShadeThousandth &&
            !accent4Tint.HasLuminanceModulationThousandth &&
            !accent4Tint.HasLuminanceOffsetThousandth &&
            !accent4Tint.HasAlphaModulationThousandth &&
            !accent4Tint.HasAlphaOffsetThousandth &&
            !accent4Tint.HasSaturationModulationThousandth &&
            !accent4Tint.HasSaturationOffsetThousandth &&
            !accent4Tint.HasRedModulationThousandth &&
            !accent4Tint.HasRedOffsetThousandth &&
            !accent4Tint.HasGreenModulationThousandth &&
            !accent4Tint.HasGreenOffsetThousandth &&
            !accent4Tint.HasBlueModulationThousandth &&
            !accent4Tint.HasBlueOffsetThousandth &&
            !accent4Tint.HasHueModulationThousandth &&
            !accent4Tint.HasHueOffsetAngleThousandth &&
            !accent4Tint.HasGray && !accent4Tint.HasComp && !accent4Tint.HasInv &&
            !accent4Tint.HasGamma && !accent4Tint.HasInvGamma &&
            accent4Tint.TintThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4Tint(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 tint.",
                    PartPath(themePart));
            var requested = (long)accent4Tint.TintThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var tint = color.FirstChild as A.Tint ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 tint.",
                        PartPath(themePart));
                tint.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasShadeThousandth: true } accent5Shade &&
            !accent5Shade.HasTintThousandth &&
            !accent5Shade.HasLuminanceModulationThousandth &&
            !accent5Shade.HasLuminanceOffsetThousandth &&
            !accent5Shade.HasAlphaModulationThousandth &&
            !accent5Shade.HasAlphaOffsetThousandth &&
            !accent5Shade.HasSaturationModulationThousandth &&
            !accent5Shade.HasSaturationOffsetThousandth &&
            !accent5Shade.HasRedModulationThousandth &&
            !accent5Shade.HasRedOffsetThousandth &&
            !accent5Shade.HasGreenModulationThousandth &&
            !accent5Shade.HasGreenOffsetThousandth &&
            !accent5Shade.HasBlueModulationThousandth &&
            !accent5Shade.HasBlueOffsetThousandth &&
            !accent5Shade.HasHueModulationThousandth &&
            !accent5Shade.HasHueOffsetAngleThousandth &&
            !accent5Shade.HasGray && !accent5Shade.HasComp && !accent5Shade.HasInv &&
            !accent5Shade.HasGamma && !accent5Shade.HasInvGamma &&
            accent5Shade.ShadeThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5Shade(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 shade.",
                    PartPath(themePart));
            var requested = (long)accent5Shade.ShadeThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var shade = color.FirstChild as A.Shade ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 shade.",
                        PartPath(themePart));
                shade.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasShadeThousandth: true } accent4Shade &&
            !accent4Shade.HasTintThousandth &&
            !accent4Shade.HasLuminanceModulationThousandth &&
            !accent4Shade.HasLuminanceOffsetThousandth &&
            !accent4Shade.HasAlphaModulationThousandth &&
            !accent4Shade.HasAlphaOffsetThousandth &&
            !accent4Shade.HasSaturationModulationThousandth &&
            !accent4Shade.HasSaturationOffsetThousandth &&
            !accent4Shade.HasRedModulationThousandth &&
            !accent4Shade.HasRedOffsetThousandth &&
            !accent4Shade.HasGreenModulationThousandth &&
            !accent4Shade.HasGreenOffsetThousandth &&
            !accent4Shade.HasBlueModulationThousandth &&
            !accent4Shade.HasBlueOffsetThousandth &&
            !accent4Shade.HasHueModulationThousandth &&
            !accent4Shade.HasHueOffsetAngleThousandth &&
            !accent4Shade.HasGray && !accent4Shade.HasComp && !accent4Shade.HasInv &&
            !accent4Shade.HasGamma && !accent4Shade.HasInvGamma &&
            accent4Shade.ShadeThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4Shade(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 shade.",
                    PartPath(themePart));
            var requested = (long)accent4Shade.ShadeThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var shade = color.FirstChild as A.Shade ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 shade.",
                        PartPath(themePart));
                shade.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasLuminanceModulationThousandth: true } accent5LumMod &&
            !accent5LumMod.HasTintThousandth &&
            !accent5LumMod.HasShadeThousandth &&
            !accent5LumMod.HasLuminanceOffsetThousandth &&
            !accent5LumMod.HasAlphaModulationThousandth &&
            !accent5LumMod.HasAlphaOffsetThousandth &&
            !accent5LumMod.HasSaturationModulationThousandth &&
            !accent5LumMod.HasSaturationOffsetThousandth &&
            !accent5LumMod.HasRedModulationThousandth &&
            !accent5LumMod.HasRedOffsetThousandth &&
            !accent5LumMod.HasGreenModulationThousandth &&
            !accent5LumMod.HasGreenOffsetThousandth &&
            !accent5LumMod.HasBlueModulationThousandth &&
            !accent5LumMod.HasBlueOffsetThousandth &&
            !accent5LumMod.HasHueModulationThousandth &&
            !accent5LumMod.HasHueOffsetAngleThousandth &&
            !accent5LumMod.HasGray && !accent5LumMod.HasComp && !accent5LumMod.HasInv &&
            !accent5LumMod.HasGamma && !accent5LumMod.HasInvGamma &&
            accent5LumMod.LuminanceModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5LumMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 lumMod.",
                    PartPath(themePart));
            var requested = (long)accent5LumMod.LuminanceModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var lumMod = color.FirstChild as A.LuminanceModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 lumMod.",
                        PartPath(themePart));
                lumMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent6", HasLuminanceModulationThousandth: true } accent6LumMod &&
            !accent6LumMod.HasTintThousandth &&
            !accent6LumMod.HasShadeThousandth &&
            !accent6LumMod.HasLuminanceOffsetThousandth &&
            !accent6LumMod.HasAlphaModulationThousandth &&
            !accent6LumMod.HasAlphaOffsetThousandth &&
            !accent6LumMod.HasSaturationModulationThousandth &&
            !accent6LumMod.HasSaturationOffsetThousandth &&
            !accent6LumMod.HasRedModulationThousandth &&
            !accent6LumMod.HasRedOffsetThousandth &&
            !accent6LumMod.HasGreenModulationThousandth &&
            !accent6LumMod.HasGreenOffsetThousandth &&
            !accent6LumMod.HasBlueModulationThousandth &&
            !accent6LumMod.HasBlueOffsetThousandth &&
            !accent6LumMod.HasHueModulationThousandth &&
            !accent6LumMod.HasHueOffsetAngleThousandth &&
            !accent6LumMod.HasGray && !accent6LumMod.HasComp && !accent6LumMod.HasInv &&
            !accent6LumMod.HasGamma && !accent6LumMod.HasInvGamma &&
            accent6LumMod.LuminanceModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent6LumMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent6 lumMod.",
                    PartPath(themePart));
            var requested = (long)accent6LumMod.LuminanceModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent6Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 color.",
                        PartPath(themePart));
                var lumMod = color.FirstChild as A.LuminanceModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent6 lumMod.",
                        PartPath(themePart));
                lumMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasLuminanceModulationThousandth: true } accent4LumMod &&
            !accent4LumMod.HasTintThousandth &&
            !accent4LumMod.HasShadeThousandth &&
            !accent4LumMod.HasLuminanceOffsetThousandth &&
            !accent4LumMod.HasAlphaModulationThousandth &&
            !accent4LumMod.HasAlphaOffsetThousandth &&
            !accent4LumMod.HasSaturationModulationThousandth &&
            !accent4LumMod.HasSaturationOffsetThousandth &&
            !accent4LumMod.HasRedModulationThousandth &&
            !accent4LumMod.HasRedOffsetThousandth &&
            !accent4LumMod.HasGreenModulationThousandth &&
            !accent4LumMod.HasGreenOffsetThousandth &&
            !accent4LumMod.HasBlueModulationThousandth &&
            !accent4LumMod.HasBlueOffsetThousandth &&
            !accent4LumMod.HasHueModulationThousandth &&
            !accent4LumMod.HasHueOffsetAngleThousandth &&
            !accent4LumMod.HasGray && !accent4LumMod.HasComp && !accent4LumMod.HasInv &&
            !accent4LumMod.HasGamma && !accent4LumMod.HasInvGamma &&
            accent4LumMod.LuminanceModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4LumMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 lumMod.",
                    PartPath(themePart));
            var requested = (long)accent4LumMod.LuminanceModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var lumMod = color.FirstChild as A.LuminanceModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 lumMod.",
                        PartPath(themePart));
                lumMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasLuminanceOffsetThousandth: true } accent5LumOff &&
            !accent5LumOff.HasTintThousandth &&
            !accent5LumOff.HasShadeThousandth &&
            !accent5LumOff.HasLuminanceModulationThousandth &&
            !accent5LumOff.HasAlphaModulationThousandth &&
            !accent5LumOff.HasAlphaOffsetThousandth &&
            !accent5LumOff.HasSaturationModulationThousandth &&
            !accent5LumOff.HasSaturationOffsetThousandth &&
            !accent5LumOff.HasRedModulationThousandth &&
            !accent5LumOff.HasRedOffsetThousandth &&
            !accent5LumOff.HasGreenModulationThousandth &&
            !accent5LumOff.HasGreenOffsetThousandth &&
            !accent5LumOff.HasBlueModulationThousandth &&
            !accent5LumOff.HasBlueOffsetThousandth &&
            !accent5LumOff.HasHueModulationThousandth &&
            !accent5LumOff.HasHueOffsetAngleThousandth &&
            !accent5LumOff.HasGray && !accent5LumOff.HasComp && !accent5LumOff.HasInv &&
            !accent5LumOff.HasGamma && !accent5LumOff.HasInvGamma &&
            accent5LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent5LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5LumOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 lumOff.",
                    PartPath(themePart));
            var requested = (long)accent5LumOff.LuminanceOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var lumOff = color.FirstChild as A.LuminanceOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 lumOff.",
                        PartPath(themePart));
                lumOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasAlphaModulationThousandth: true } accent5AlphaMod &&
            !accent5AlphaMod.HasTintThousandth &&
            !accent5AlphaMod.HasShadeThousandth &&
            !accent5AlphaMod.HasLuminanceModulationThousandth &&
            !accent5AlphaMod.HasLuminanceOffsetThousandth &&
            !accent5AlphaMod.HasAlphaOffsetThousandth &&
            !accent5AlphaMod.HasSaturationModulationThousandth &&
            !accent5AlphaMod.HasSaturationOffsetThousandth &&
            !accent5AlphaMod.HasRedModulationThousandth &&
            !accent5AlphaMod.HasRedOffsetThousandth &&
            !accent5AlphaMod.HasGreenModulationThousandth &&
            !accent5AlphaMod.HasGreenOffsetThousandth &&
            !accent5AlphaMod.HasBlueModulationThousandth &&
            !accent5AlphaMod.HasBlueOffsetThousandth &&
            !accent5AlphaMod.HasHueModulationThousandth &&
            !accent5AlphaMod.HasHueOffsetAngleThousandth &&
            !accent5AlphaMod.HasGray && !accent5AlphaMod.HasComp && !accent5AlphaMod.HasInv &&
            !accent5AlphaMod.HasGamma && !accent5AlphaMod.HasInvGamma &&
            accent5AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5AlphaMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 alphaMod.",
                    PartPath(themePart));
            var requested = (long)accent5AlphaMod.AlphaModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var alphaMod = color.FirstChild as A.AlphaModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 alphaMod.",
                        PartPath(themePart));
                alphaMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasAlphaOffsetThousandth: true } accent5AlphaOff &&
            !accent5AlphaOff.HasTintThousandth &&
            !accent5AlphaOff.HasShadeThousandth &&
            !accent5AlphaOff.HasLuminanceModulationThousandth &&
            !accent5AlphaOff.HasLuminanceOffsetThousandth &&
            !accent5AlphaOff.HasAlphaModulationThousandth &&
            !accent5AlphaOff.HasSaturationModulationThousandth &&
            !accent5AlphaOff.HasSaturationOffsetThousandth &&
            !accent5AlphaOff.HasRedModulationThousandth &&
            !accent5AlphaOff.HasRedOffsetThousandth &&
            !accent5AlphaOff.HasGreenModulationThousandth &&
            !accent5AlphaOff.HasGreenOffsetThousandth &&
            !accent5AlphaOff.HasBlueModulationThousandth &&
            !accent5AlphaOff.HasBlueOffsetThousandth &&
            !accent5AlphaOff.HasHueModulationThousandth &&
            !accent5AlphaOff.HasHueOffsetAngleThousandth &&
            !accent5AlphaOff.HasGray && !accent5AlphaOff.HasComp && !accent5AlphaOff.HasInv &&
            !accent5AlphaOff.HasGamma && !accent5AlphaOff.HasInvGamma &&
            accent5AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent5AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5AlphaOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 alphaOff.",
                    PartPath(themePart));
            var requested = (long)accent5AlphaOff.AlphaOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var alphaOff = color.FirstChild as A.AlphaOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 alphaOff.",
                        PartPath(themePart));
                alphaOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasSaturationModulationThousandth: true } accent5SatMod &&
            !accent5SatMod.HasTintThousandth &&
            !accent5SatMod.HasShadeThousandth &&
            !accent5SatMod.HasLuminanceModulationThousandth &&
            !accent5SatMod.HasLuminanceOffsetThousandth &&
            !accent5SatMod.HasAlphaModulationThousandth &&
            !accent5SatMod.HasAlphaOffsetThousandth &&
            !accent5SatMod.HasSaturationOffsetThousandth &&
            !accent5SatMod.HasRedModulationThousandth &&
            !accent5SatMod.HasRedOffsetThousandth &&
            !accent5SatMod.HasGreenModulationThousandth &&
            !accent5SatMod.HasGreenOffsetThousandth &&
            !accent5SatMod.HasBlueModulationThousandth &&
            !accent5SatMod.HasBlueOffsetThousandth &&
            !accent5SatMod.HasHueModulationThousandth &&
            !accent5SatMod.HasHueOffsetAngleThousandth &&
            !accent5SatMod.HasGray && !accent5SatMod.HasComp && !accent5SatMod.HasInv &&
            !accent5SatMod.HasGamma && !accent5SatMod.HasInvGamma &&
            accent5SatMod.SaturationModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5SatMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 satMod.",
                    PartPath(themePart));
            var requested = (long)accent5SatMod.SaturationModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var satMod = color.FirstChild as A.SaturationModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 satMod.",
                        PartPath(themePart));
                satMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasSaturationOffsetThousandth: true } accent5SatOff &&
            !accent5SatOff.HasTintThousandth &&
            !accent5SatOff.HasShadeThousandth &&
            !accent5SatOff.HasLuminanceModulationThousandth &&
            !accent5SatOff.HasLuminanceOffsetThousandth &&
            !accent5SatOff.HasAlphaModulationThousandth &&
            !accent5SatOff.HasAlphaOffsetThousandth &&
            !accent5SatOff.HasSaturationModulationThousandth &&
            !accent5SatOff.HasRedModulationThousandth &&
            !accent5SatOff.HasRedOffsetThousandth &&
            !accent5SatOff.HasGreenModulationThousandth &&
            !accent5SatOff.HasGreenOffsetThousandth &&
            !accent5SatOff.HasBlueModulationThousandth &&
            !accent5SatOff.HasBlueOffsetThousandth &&
            !accent5SatOff.HasHueModulationThousandth &&
            !accent5SatOff.HasHueOffsetAngleThousandth &&
            !accent5SatOff.HasGray && !accent5SatOff.HasComp && !accent5SatOff.HasInv &&
            !accent5SatOff.HasGamma && !accent5SatOff.HasInvGamma &&
            accent5SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent5SatOff.SaturationOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5SatOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 satOff.",
                    PartPath(themePart));
            var requested = (long)accent5SatOff.SaturationOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var satOff = color.FirstChild as A.SaturationOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 satOff.",
                        PartPath(themePart));
                satOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasRedModulationThousandth: true } accent5RedMod &&
            !accent5RedMod.HasTintThousandth &&
            !accent5RedMod.HasShadeThousandth &&
            !accent5RedMod.HasLuminanceModulationThousandth &&
            !accent5RedMod.HasLuminanceOffsetThousandth &&
            !accent5RedMod.HasAlphaModulationThousandth &&
            !accent5RedMod.HasAlphaOffsetThousandth &&
            !accent5RedMod.HasSaturationModulationThousandth &&
            !accent5RedMod.HasSaturationOffsetThousandth &&
            !accent5RedMod.HasRedOffsetThousandth &&
            !accent5RedMod.HasGreenModulationThousandth &&
            !accent5RedMod.HasGreenOffsetThousandth &&
            !accent5RedMod.HasBlueModulationThousandth &&
            !accent5RedMod.HasBlueOffsetThousandth &&
            !accent5RedMod.HasHueModulationThousandth &&
            !accent5RedMod.HasHueOffsetAngleThousandth &&
            !accent5RedMod.HasGray && !accent5RedMod.HasComp && !accent5RedMod.HasInv &&
            !accent5RedMod.HasGamma && !accent5RedMod.HasInvGamma &&
            accent5RedMod.RedModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5RedMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 redMod.",
                    PartPath(themePart));
            var requested = (long)accent5RedMod.RedModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var redMod = color.FirstChild as A.RedModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 redMod.",
                        PartPath(themePart));
                redMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasRedOffsetThousandth: true } accent5RedOff &&
            !accent5RedOff.HasTintThousandth &&
            !accent5RedOff.HasShadeThousandth &&
            !accent5RedOff.HasLuminanceModulationThousandth &&
            !accent5RedOff.HasLuminanceOffsetThousandth &&
            !accent5RedOff.HasAlphaModulationThousandth &&
            !accent5RedOff.HasAlphaOffsetThousandth &&
            !accent5RedOff.HasSaturationModulationThousandth &&
            !accent5RedOff.HasSaturationOffsetThousandth &&
            !accent5RedOff.HasRedModulationThousandth &&
            !accent5RedOff.HasGreenModulationThousandth &&
            !accent5RedOff.HasGreenOffsetThousandth &&
            !accent5RedOff.HasBlueModulationThousandth &&
            !accent5RedOff.HasBlueOffsetThousandth &&
            !accent5RedOff.HasHueModulationThousandth &&
            !accent5RedOff.HasHueOffsetAngleThousandth &&
            !accent5RedOff.HasGray && !accent5RedOff.HasComp && !accent5RedOff.HasInv &&
            !accent5RedOff.HasGamma && !accent5RedOff.HasInvGamma &&
            accent5RedOff.RedOffsetThousandth >= -100_000 &&
            accent5RedOff.RedOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5RedOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 redOff.",
                    PartPath(themePart));
            var requested = (long)accent5RedOff.RedOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var redOff = color.FirstChild as A.RedOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 redOff.",
                        PartPath(themePart));
                redOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasGreenModulationThousandth: true } accent5GreenMod &&
            !accent5GreenMod.HasTintThousandth &&
            !accent5GreenMod.HasShadeThousandth &&
            !accent5GreenMod.HasLuminanceModulationThousandth &&
            !accent5GreenMod.HasLuminanceOffsetThousandth &&
            !accent5GreenMod.HasAlphaModulationThousandth &&
            !accent5GreenMod.HasAlphaOffsetThousandth &&
            !accent5GreenMod.HasSaturationModulationThousandth &&
            !accent5GreenMod.HasSaturationOffsetThousandth &&
            !accent5GreenMod.HasRedOffsetThousandth &&
            !accent5GreenMod.HasGreenOffsetThousandth &&
            !accent5GreenMod.HasBlueModulationThousandth &&
            !accent5GreenMod.HasBlueOffsetThousandth &&
            !accent5GreenMod.HasHueModulationThousandth &&
            !accent5GreenMod.HasHueOffsetAngleThousandth &&
            !accent5GreenMod.HasGray && !accent5GreenMod.HasComp && !accent5GreenMod.HasInv &&
            !accent5GreenMod.HasGamma && !accent5GreenMod.HasInvGamma &&
            accent5GreenMod.GreenModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5GreenMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 greenMod.",
                    PartPath(themePart));
            var requested = (long)accent5GreenMod.GreenModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var greenMod = color.FirstChild as A.GreenModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 greenMod.",
                        PartPath(themePart));
                greenMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasGreenOffsetThousandth: true } accent5GreenOff &&
            !accent5GreenOff.HasTintThousandth &&
            !accent5GreenOff.HasShadeThousandth &&
            !accent5GreenOff.HasLuminanceModulationThousandth &&
            !accent5GreenOff.HasLuminanceOffsetThousandth &&
            !accent5GreenOff.HasAlphaModulationThousandth &&
            !accent5GreenOff.HasAlphaOffsetThousandth &&
            !accent5GreenOff.HasSaturationModulationThousandth &&
            !accent5GreenOff.HasSaturationOffsetThousandth &&
            !accent5GreenOff.HasRedOffsetThousandth &&
            !accent5GreenOff.HasGreenModulationThousandth &&
            !accent5GreenOff.HasBlueModulationThousandth &&
            !accent5GreenOff.HasBlueOffsetThousandth &&
            !accent5GreenOff.HasHueModulationThousandth &&
            !accent5GreenOff.HasHueOffsetAngleThousandth &&
            !accent5GreenOff.HasGray && !accent5GreenOff.HasComp && !accent5GreenOff.HasInv &&
            !accent5GreenOff.HasGamma && !accent5GreenOff.HasInvGamma &&
            accent5GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent5GreenOff.GreenOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5GreenOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 greenOff.",
                    PartPath(themePart));
            var requested = (long)accent5GreenOff.GreenOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var greenOff = color.FirstChild as A.GreenOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 greenOff.",
                        PartPath(themePart));
                greenOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasBlueModulationThousandth: true } accent5BlueMod &&
            !accent5BlueMod.HasTintThousandth &&
            !accent5BlueMod.HasShadeThousandth &&
            !accent5BlueMod.HasLuminanceModulationThousandth &&
            !accent5BlueMod.HasLuminanceOffsetThousandth &&
            !accent5BlueMod.HasAlphaModulationThousandth &&
            !accent5BlueMod.HasAlphaOffsetThousandth &&
            !accent5BlueMod.HasSaturationModulationThousandth &&
            !accent5BlueMod.HasSaturationOffsetThousandth &&
            !accent5BlueMod.HasRedOffsetThousandth &&
            !accent5BlueMod.HasGreenOffsetThousandth &&
            !accent5BlueMod.HasGreenModulationThousandth &&
            !accent5BlueMod.HasBlueOffsetThousandth &&
            !accent5BlueMod.HasHueModulationThousandth &&
            !accent5BlueMod.HasHueOffsetAngleThousandth &&
            !accent5BlueMod.HasGray && !accent5BlueMod.HasComp && !accent5BlueMod.HasInv &&
            !accent5BlueMod.HasGamma && !accent5BlueMod.HasInvGamma &&
            accent5BlueMod.BlueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5BlueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 blueMod.",
                    PartPath(themePart));
            var requested = (long)accent5BlueMod.BlueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var blueMod = color.FirstChild as A.BlueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 blueMod.",
                        PartPath(themePart));
                blueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasBlueOffsetThousandth: true } accent5BlueOff &&
            !accent5BlueOff.HasTintThousandth &&
            !accent5BlueOff.HasShadeThousandth &&
            !accent5BlueOff.HasLuminanceModulationThousandth &&
            !accent5BlueOff.HasLuminanceOffsetThousandth &&
            !accent5BlueOff.HasAlphaModulationThousandth &&
            !accent5BlueOff.HasAlphaOffsetThousandth &&
            !accent5BlueOff.HasSaturationModulationThousandth &&
            !accent5BlueOff.HasSaturationOffsetThousandth &&
            !accent5BlueOff.HasRedOffsetThousandth &&
            !accent5BlueOff.HasGreenModulationThousandth &&
            !accent5BlueOff.HasBlueModulationThousandth &&
            !accent5BlueOff.HasGreenOffsetThousandth &&
            !accent5BlueOff.HasHueModulationThousandth &&
            !accent5BlueOff.HasHueOffsetAngleThousandth &&
            !accent5BlueOff.HasGray && !accent5BlueOff.HasComp && !accent5BlueOff.HasInv &&
            !accent5BlueOff.HasGamma && !accent5BlueOff.HasInvGamma &&
            accent5BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent5BlueOff.BlueOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5BlueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 blueOff.",
                    PartPath(themePart));
            var requested = (long)accent5BlueOff.BlueOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var blueOff = color.FirstChild as A.BlueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 blueOff.",
                        PartPath(themePart));
                blueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasHueModulationThousandth: true } accent5HueMod &&
            !accent5HueMod.HasTintThousandth &&
            !accent5HueMod.HasShadeThousandth &&
            !accent5HueMod.HasLuminanceModulationThousandth &&
            !accent5HueMod.HasLuminanceOffsetThousandth &&
            !accent5HueMod.HasAlphaModulationThousandth &&
            !accent5HueMod.HasAlphaOffsetThousandth &&
            !accent5HueMod.HasSaturationModulationThousandth &&
            !accent5HueMod.HasSaturationOffsetThousandth &&
            !accent5HueMod.HasRedModulationThousandth &&
            !accent5HueMod.HasRedOffsetThousandth &&
            !accent5HueMod.HasGreenModulationThousandth &&
            !accent5HueMod.HasGreenOffsetThousandth &&
            !accent5HueMod.HasBlueModulationThousandth &&
            !accent5HueMod.HasHueOffsetAngleThousandth &&
            !accent5HueMod.HasGray && !accent5HueMod.HasComp && !accent5HueMod.HasInv &&
            !accent5HueMod.HasGamma && !accent5HueMod.HasInvGamma &&
            accent5HueMod.HueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5HueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 hueMod.",
                    PartPath(themePart));
            var requested = (long)accent5HueMod.HueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var hueMod = color.FirstChild as A.HueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 hueMod.",
                        PartPath(themePart));
                hueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent5", HasHueOffsetAngleThousandth: true } accent5HueOff &&
            !accent5HueOff.HasTintThousandth &&
            !accent5HueOff.HasShadeThousandth &&
            !accent5HueOff.HasLuminanceModulationThousandth &&
            !accent5HueOff.HasLuminanceOffsetThousandth &&
            !accent5HueOff.HasAlphaModulationThousandth &&
            !accent5HueOff.HasAlphaOffsetThousandth &&
            !accent5HueOff.HasSaturationModulationThousandth &&
            !accent5HueOff.HasSaturationOffsetThousandth &&
            !accent5HueOff.HasRedModulationThousandth &&
            !accent5HueOff.HasRedOffsetThousandth &&
            !accent5HueOff.HasGreenModulationThousandth &&
            !accent5HueOff.HasGreenOffsetThousandth &&
            !accent5HueOff.HasBlueModulationThousandth &&
            !accent5HueOff.HasHueModulationThousandth &&
            !accent5HueOff.HasGray && !accent5HueOff.HasComp && !accent5HueOff.HasInv &&
            !accent5HueOff.HasGamma && !accent5HueOff.HasInvGamma &&
            accent5HueOff.HueOffsetAngleThousandth >= -21_600_000 &&
            accent5HueOff.HueOffsetAngleThousandth <= 21_600_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent5HueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent5 hueOff.",
                    PartPath(themePart));
            var requested = (long)accent5HueOff.HueOffsetAngleThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent5Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 color.",
                        PartPath(themePart));
                var hueOff = color.FirstChild as A.HueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent5 hueOff.",
                        PartPath(themePart));
                hueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasLuminanceOffsetThousandth: true } accent4LumOff &&
            !accent4LumOff.HasTintThousandth &&
            !accent4LumOff.HasShadeThousandth &&
            !accent4LumOff.HasLuminanceModulationThousandth &&
            !accent4LumOff.HasAlphaModulationThousandth &&
            !accent4LumOff.HasAlphaOffsetThousandth &&
            !accent4LumOff.HasSaturationModulationThousandth &&
            !accent4LumOff.HasSaturationOffsetThousandth &&
            !accent4LumOff.HasRedModulationThousandth &&
            !accent4LumOff.HasRedOffsetThousandth &&
            !accent4LumOff.HasGreenModulationThousandth &&
            !accent4LumOff.HasGreenOffsetThousandth &&
            !accent4LumOff.HasBlueModulationThousandth &&
            !accent4LumOff.HasBlueOffsetThousandth &&
            !accent4LumOff.HasHueModulationThousandth &&
            !accent4LumOff.HasHueOffsetAngleThousandth &&
            !accent4LumOff.HasGray && !accent4LumOff.HasComp && !accent4LumOff.HasInv &&
            !accent4LumOff.HasGamma && !accent4LumOff.HasInvGamma &&
            accent4LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent4LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4LumOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 lumOff.",
                    PartPath(themePart));
            var requested = (long)accent4LumOff.LuminanceOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var lumOff = color.FirstChild as A.LuminanceOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 lumOff.",
                        PartPath(themePart));
                lumOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasAlphaModulationThousandth: true } accent4AlphaMod &&
            !accent4AlphaMod.HasTintThousandth &&
            !accent4AlphaMod.HasShadeThousandth &&
            !accent4AlphaMod.HasLuminanceModulationThousandth &&
            !accent4AlphaMod.HasLuminanceOffsetThousandth &&
            !accent4AlphaMod.HasAlphaOffsetThousandth &&
            !accent4AlphaMod.HasSaturationModulationThousandth &&
            !accent4AlphaMod.HasSaturationOffsetThousandth &&
            !accent4AlphaMod.HasRedModulationThousandth &&
            !accent4AlphaMod.HasRedOffsetThousandth &&
            !accent4AlphaMod.HasGreenModulationThousandth &&
            !accent4AlphaMod.HasGreenOffsetThousandth &&
            !accent4AlphaMod.HasBlueModulationThousandth &&
            !accent4AlphaMod.HasBlueOffsetThousandth &&
            !accent4AlphaMod.HasHueModulationThousandth &&
            !accent4AlphaMod.HasHueOffsetAngleThousandth &&
            !accent4AlphaMod.HasGray && !accent4AlphaMod.HasComp && !accent4AlphaMod.HasInv &&
            !accent4AlphaMod.HasGamma && !accent4AlphaMod.HasInvGamma &&
            accent4AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4AlphaMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 alphaMod.",
                    PartPath(themePart));
            var requested = (long)accent4AlphaMod.AlphaModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var alphaMod = color.FirstChild as A.AlphaModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 alphaMod.",
                        PartPath(themePart));
                alphaMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasAlphaOffsetThousandth: true } accent4AlphaOff &&
            !accent4AlphaOff.HasTintThousandth &&
            !accent4AlphaOff.HasShadeThousandth &&
            !accent4AlphaOff.HasLuminanceModulationThousandth &&
            !accent4AlphaOff.HasLuminanceOffsetThousandth &&
            !accent4AlphaOff.HasAlphaModulationThousandth &&
            !accent4AlphaOff.HasSaturationModulationThousandth &&
            !accent4AlphaOff.HasSaturationOffsetThousandth &&
            !accent4AlphaOff.HasRedModulationThousandth &&
            !accent4AlphaOff.HasRedOffsetThousandth &&
            !accent4AlphaOff.HasGreenModulationThousandth &&
            !accent4AlphaOff.HasGreenOffsetThousandth &&
            !accent4AlphaOff.HasBlueModulationThousandth &&
            !accent4AlphaOff.HasBlueOffsetThousandth &&
            !accent4AlphaOff.HasHueModulationThousandth &&
            !accent4AlphaOff.HasHueOffsetAngleThousandth &&
            !accent4AlphaOff.HasGray && !accent4AlphaOff.HasComp && !accent4AlphaOff.HasInv &&
            !accent4AlphaOff.HasGamma && !accent4AlphaOff.HasInvGamma &&
            accent4AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent4AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4AlphaOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 alphaOff.",
                    PartPath(themePart));
            var requested = (long)accent4AlphaOff.AlphaOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var alphaOff = color.FirstChild as A.AlphaOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 alphaOff.",
                        PartPath(themePart));
                alphaOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasSaturationModulationThousandth: true } accent4SatMod &&
            !accent4SatMod.HasTintThousandth &&
            !accent4SatMod.HasShadeThousandth &&
            !accent4SatMod.HasLuminanceModulationThousandth &&
            !accent4SatMod.HasLuminanceOffsetThousandth &&
            !accent4SatMod.HasAlphaModulationThousandth &&
            !accent4SatMod.HasAlphaOffsetThousandth &&
            !accent4SatMod.HasSaturationOffsetThousandth &&
            !accent4SatMod.HasRedModulationThousandth &&
            !accent4SatMod.HasRedOffsetThousandth &&
            !accent4SatMod.HasGreenModulationThousandth &&
            !accent4SatMod.HasGreenOffsetThousandth &&
            !accent4SatMod.HasBlueModulationThousandth &&
            !accent4SatMod.HasBlueOffsetThousandth &&
            !accent4SatMod.HasHueModulationThousandth &&
            !accent4SatMod.HasHueOffsetAngleThousandth &&
            !accent4SatMod.HasGray && !accent4SatMod.HasComp && !accent4SatMod.HasInv &&
            !accent4SatMod.HasGamma && !accent4SatMod.HasInvGamma &&
            accent4SatMod.SaturationModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4SatMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 satMod.",
                    PartPath(themePart));
            var requested = (long)accent4SatMod.SaturationModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var satMod = color.FirstChild as A.SaturationModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 satMod.",
                        PartPath(themePart));
                satMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasSaturationOffsetThousandth: true } accent4SatOff &&
            !accent4SatOff.HasTintThousandth &&
            !accent4SatOff.HasShadeThousandth &&
            !accent4SatOff.HasLuminanceModulationThousandth &&
            !accent4SatOff.HasLuminanceOffsetThousandth &&
            !accent4SatOff.HasAlphaModulationThousandth &&
            !accent4SatOff.HasAlphaOffsetThousandth &&
            !accent4SatOff.HasSaturationModulationThousandth &&
            !accent4SatOff.HasRedModulationThousandth &&
            !accent4SatOff.HasRedOffsetThousandth &&
            !accent4SatOff.HasGreenModulationThousandth &&
            !accent4SatOff.HasGreenOffsetThousandth &&
            !accent4SatOff.HasBlueModulationThousandth &&
            !accent4SatOff.HasBlueOffsetThousandth &&
            !accent4SatOff.HasHueModulationThousandth &&
            !accent4SatOff.HasHueOffsetAngleThousandth &&
            !accent4SatOff.HasGray && !accent4SatOff.HasComp && !accent4SatOff.HasInv &&
            !accent4SatOff.HasGamma && !accent4SatOff.HasInvGamma &&
            accent4SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent4SatOff.SaturationOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4SatOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 satOff.",
                    PartPath(themePart));
            var requested = (long)accent4SatOff.SaturationOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var satOff = color.FirstChild as A.SaturationOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 satOff.",
                        PartPath(themePart));
                satOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasRedModulationThousandth: true } accent4RedMod &&
            !accent4RedMod.HasTintThousandth &&
            !accent4RedMod.HasShadeThousandth &&
            !accent4RedMod.HasLuminanceModulationThousandth &&
            !accent4RedMod.HasLuminanceOffsetThousandth &&
            !accent4RedMod.HasAlphaModulationThousandth &&
            !accent4RedMod.HasAlphaOffsetThousandth &&
            !accent4RedMod.HasSaturationModulationThousandth &&
            !accent4RedMod.HasSaturationOffsetThousandth &&
            !accent4RedMod.HasRedOffsetThousandth &&
            !accent4RedMod.HasGreenModulationThousandth &&
            !accent4RedMod.HasGreenOffsetThousandth &&
            !accent4RedMod.HasBlueModulationThousandth &&
            !accent4RedMod.HasBlueOffsetThousandth &&
            !accent4RedMod.HasHueModulationThousandth &&
            !accent4RedMod.HasHueOffsetAngleThousandth &&
            !accent4RedMod.HasGray && !accent4RedMod.HasComp && !accent4RedMod.HasInv &&
            !accent4RedMod.HasGamma && !accent4RedMod.HasInvGamma &&
            accent4RedMod.RedModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4RedMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 redMod.",
                    PartPath(themePart));
            var requested = (long)accent4RedMod.RedModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var redMod = color.FirstChild as A.RedModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 redMod.",
                        PartPath(themePart));
                redMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasGreenModulationThousandth: true } accent4GreenMod &&
            !accent4GreenMod.HasTintThousandth &&
            !accent4GreenMod.HasShadeThousandth &&
            !accent4GreenMod.HasLuminanceModulationThousandth &&
            !accent4GreenMod.HasLuminanceOffsetThousandth &&
            !accent4GreenMod.HasAlphaModulationThousandth &&
            !accent4GreenMod.HasAlphaOffsetThousandth &&
            !accent4GreenMod.HasSaturationModulationThousandth &&
            !accent4GreenMod.HasSaturationOffsetThousandth &&
            !accent4GreenMod.HasRedOffsetThousandth &&
            !accent4GreenMod.HasGreenOffsetThousandth &&
            !accent4GreenMod.HasBlueModulationThousandth &&
            !accent4GreenMod.HasBlueOffsetThousandth &&
            !accent4GreenMod.HasHueModulationThousandth &&
            !accent4GreenMod.HasHueOffsetAngleThousandth &&
            !accent4GreenMod.HasGray && !accent4GreenMod.HasComp && !accent4GreenMod.HasInv &&
            !accent4GreenMod.HasGamma && !accent4GreenMod.HasInvGamma &&
            accent4GreenMod.GreenModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4GreenMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 greenMod.",
                    PartPath(themePart));
            var requested = (long)accent4GreenMod.GreenModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var greenMod = color.FirstChild as A.GreenModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 greenMod.",
                        PartPath(themePart));
                greenMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasBlueModulationThousandth: true } accent4BlueMod &&
            !accent4BlueMod.HasTintThousandth &&
            !accent4BlueMod.HasShadeThousandth &&
            !accent4BlueMod.HasLuminanceModulationThousandth &&
            !accent4BlueMod.HasLuminanceOffsetThousandth &&
            !accent4BlueMod.HasAlphaModulationThousandth &&
            !accent4BlueMod.HasAlphaOffsetThousandth &&
            !accent4BlueMod.HasSaturationModulationThousandth &&
            !accent4BlueMod.HasSaturationOffsetThousandth &&
            !accent4BlueMod.HasRedOffsetThousandth &&
            !accent4BlueMod.HasGreenOffsetThousandth &&
            !accent4BlueMod.HasGreenModulationThousandth &&
            !accent4BlueMod.HasBlueOffsetThousandth &&
            !accent4BlueMod.HasHueModulationThousandth &&
            !accent4BlueMod.HasHueOffsetAngleThousandth &&
            !accent4BlueMod.HasGray && !accent4BlueMod.HasComp && !accent4BlueMod.HasInv &&
            !accent4BlueMod.HasGamma && !accent4BlueMod.HasInvGamma &&
            accent4BlueMod.BlueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4BlueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 blueMod.",
                    PartPath(themePart));
            var requested = (long)accent4BlueMod.BlueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var blueMod = color.FirstChild as A.BlueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 blueMod.",
                        PartPath(themePart));
                blueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasHueModulationThousandth: true } accent4HueMod &&
            !accent4HueMod.HasTintThousandth &&
            !accent4HueMod.HasShadeThousandth &&
            !accent4HueMod.HasLuminanceModulationThousandth &&
            !accent4HueMod.HasLuminanceOffsetThousandth &&
            !accent4HueMod.HasAlphaModulationThousandth &&
            !accent4HueMod.HasAlphaOffsetThousandth &&
            !accent4HueMod.HasSaturationModulationThousandth &&
            !accent4HueMod.HasSaturationOffsetThousandth &&
            !accent4HueMod.HasRedModulationThousandth &&
            !accent4HueMod.HasRedOffsetThousandth &&
            !accent4HueMod.HasGreenModulationThousandth &&
            !accent4HueMod.HasGreenOffsetThousandth &&
            !accent4HueMod.HasBlueModulationThousandth &&
            !accent4HueMod.HasHueOffsetAngleThousandth &&
            !accent4HueMod.HasGray && !accent4HueMod.HasComp && !accent4HueMod.HasInv &&
            !accent4HueMod.HasGamma && !accent4HueMod.HasInvGamma &&
            accent4HueMod.HueModulationThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4HueMod(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 hueMod.",
                    PartPath(themePart));
            var requested = (long)accent4HueMod.HueModulationThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var hueMod = color.FirstChild as A.HueModulation ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 hueMod.",
                        PartPath(themePart));
                hueMod.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasBlueOffsetThousandth: true } accent4BlueOff &&
            !accent4BlueOff.HasTintThousandth &&
            !accent4BlueOff.HasShadeThousandth &&
            !accent4BlueOff.HasLuminanceModulationThousandth &&
            !accent4BlueOff.HasLuminanceOffsetThousandth &&
            !accent4BlueOff.HasAlphaModulationThousandth &&
            !accent4BlueOff.HasAlphaOffsetThousandth &&
            !accent4BlueOff.HasSaturationModulationThousandth &&
            !accent4BlueOff.HasSaturationOffsetThousandth &&
            !accent4BlueOff.HasRedOffsetThousandth &&
            !accent4BlueOff.HasGreenModulationThousandth &&
            !accent4BlueOff.HasBlueModulationThousandth &&
            !accent4BlueOff.HasGreenOffsetThousandth &&
            !accent4BlueOff.HasHueModulationThousandth &&
            !accent4BlueOff.HasHueOffsetAngleThousandth &&
            !accent4BlueOff.HasGray && !accent4BlueOff.HasComp && !accent4BlueOff.HasInv &&
            !accent4BlueOff.HasGamma && !accent4BlueOff.HasInvGamma &&
            accent4BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent4BlueOff.BlueOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4BlueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 blueOff.",
                    PartPath(themePart));
            var requested = (long)accent4BlueOff.BlueOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var blueOff = color.FirstChild as A.BlueOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 blueOff.",
                        PartPath(themePart));
                blueOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasGreenOffsetThousandth: true } accent4GreenOff &&
            !accent4GreenOff.HasTintThousandth &&
            !accent4GreenOff.HasShadeThousandth &&
            !accent4GreenOff.HasLuminanceModulationThousandth &&
            !accent4GreenOff.HasLuminanceOffsetThousandth &&
            !accent4GreenOff.HasAlphaModulationThousandth &&
            !accent4GreenOff.HasAlphaOffsetThousandth &&
            !accent4GreenOff.HasSaturationModulationThousandth &&
            !accent4GreenOff.HasSaturationOffsetThousandth &&
            !accent4GreenOff.HasRedOffsetThousandth &&
            !accent4GreenOff.HasGreenModulationThousandth &&
            !accent4GreenOff.HasBlueModulationThousandth &&
            !accent4GreenOff.HasBlueOffsetThousandth &&
            !accent4GreenOff.HasHueModulationThousandth &&
            !accent4GreenOff.HasHueOffsetAngleThousandth &&
            !accent4GreenOff.HasGray && !accent4GreenOff.HasComp && !accent4GreenOff.HasInv &&
            !accent4GreenOff.HasGamma && !accent4GreenOff.HasInvGamma &&
            accent4GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent4GreenOff.GreenOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4GreenOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 greenOff.",
                    PartPath(themePart));
            var requested = (long)accent4GreenOff.GreenOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var greenOff = color.FirstChild as A.GreenOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 greenOff.",
                        PartPath(themePart));
                greenOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count == 1 &&
            authoredTheme.AccentTransforms[0] is { Role: "accent4", HasRedOffsetThousandth: true } accent4RedOff &&
            !accent4RedOff.HasTintThousandth &&
            !accent4RedOff.HasShadeThousandth &&
            !accent4RedOff.HasLuminanceModulationThousandth &&
            !accent4RedOff.HasLuminanceOffsetThousandth &&
            !accent4RedOff.HasAlphaModulationThousandth &&
            !accent4RedOff.HasAlphaOffsetThousandth &&
            !accent4RedOff.HasSaturationModulationThousandth &&
            !accent4RedOff.HasSaturationOffsetThousandth &&
            !accent4RedOff.HasRedModulationThousandth &&
            !accent4RedOff.HasGreenModulationThousandth &&
            !accent4RedOff.HasGreenOffsetThousandth &&
            !accent4RedOff.HasBlueModulationThousandth &&
            !accent4RedOff.HasBlueOffsetThousandth &&
            !accent4RedOff.HasHueModulationThousandth &&
            !accent4RedOff.HasHueOffsetAngleThousandth &&
            !accent4RedOff.HasGray && !accent4RedOff.HasComp && !accent4RedOff.HasInv &&
            !accent4RedOff.HasGamma && !accent4RedOff.HasInvGamma &&
            accent4RedOff.RedOffsetThousandth >= -100_000 &&
            accent4RedOff.RedOffsetThousandth <= 100_000)
        {
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = TryReadSourceBoundThemeAccent4RedOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent4 redOff.",
                    PartPath(themePart));
            var requested = (long)accent4RedOff.RedOffsetThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent4Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 color.",
                        PartPath(themePart));
                var redOff = color.FirstChild as A.RedOffset ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent4 redOff.",
                        PartPath(themePart));
                redOff.Val = checked((int)requested);
                changed = true;
            }
        }
        else if (authoredTheme.AccentTransforms.Count > 0)
        {
            var transform = authoredTheme.AccentTransforms[0];
            var tintOnly = transform.HasTintThousandth && !transform.HasShadeThousandth;
            var shadeOnly = !transform.HasTintThousandth && transform.HasShadeThousandth;
            var lumModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                transform.HasLuminanceModulationThousandth;
            var lumOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && transform.HasLuminanceOffsetThousandth;
            var alphaModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                transform.HasAlphaModulationThousandth;
            var alphaOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && transform.HasAlphaOffsetThousandth;
            var satModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                transform.HasSaturationModulationThousandth;
            var satOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && transform.HasSaturationOffsetThousandth;
            var redModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                transform.HasRedModulationThousandth;
            var redOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && transform.HasRedOffsetThousandth;
            var greenModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                transform.HasGreenModulationThousandth;
            var greenOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                !transform.HasGreenModulationThousandth && transform.HasGreenOffsetThousandth;
            var blueModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                !transform.HasGreenModulationThousandth && !transform.HasGreenOffsetThousandth &&
                transform.HasBlueModulationThousandth;
            var blueOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                !transform.HasGreenModulationThousandth && !transform.HasGreenOffsetThousandth &&
                !transform.HasBlueModulationThousandth && transform.HasBlueOffsetThousandth;
            var hueModOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                !transform.HasGreenModulationThousandth && !transform.HasGreenOffsetThousandth &&
                !transform.HasBlueModulationThousandth && !transform.HasBlueOffsetThousandth &&
                transform.HasHueModulationThousandth;
            var hueOffOnly = !transform.HasTintThousandth && !transform.HasShadeThousandth &&
                !transform.HasLuminanceModulationThousandth && !transform.HasLuminanceOffsetThousandth &&
                !transform.HasAlphaModulationThousandth && !transform.HasAlphaOffsetThousandth &&
                !transform.HasSaturationModulationThousandth && !transform.HasSaturationOffsetThousandth &&
                !transform.HasRedModulationThousandth && !transform.HasRedOffsetThousandth &&
                !transform.HasGreenModulationThousandth && !transform.HasGreenOffsetThousandth &&
                !transform.HasBlueModulationThousandth && !transform.HasBlueOffsetThousandth &&
                !transform.HasHueModulationThousandth &&
                transform.HasHueOffsetAngleThousandth;
            var transformPath = tintOnly
                ? "$.design.theme.accentTransforms.accent1.tint"
                : shadeOnly
                    ? "$.design.theme.accentTransforms.accent1.shade"
                    : lumModOnly
                        ? "$.design.theme.accentTransforms.accent1.lumMod"
                        : lumOffOnly
                        ? "$.design.theme.accentTransforms.accent1.lumOff"
                            : alphaModOnly
                                ? "$.design.theme.accentTransforms.accent1.alphaMod"
                                : alphaOffOnly
                                    ? "$.design.theme.accentTransforms.accent1.alphaOff"
                                    : satModOnly
                                        ? "$.design.theme.accentTransforms.accent1.satMod"
                                        : satOffOnly
                                            ? "$.design.theme.accentTransforms.accent1.satOff"
                                            : redModOnly
                                                ? "$.design.theme.accentTransforms.accent1.redMod"
                                                : redOffOnly
                                                    ? "$.design.theme.accentTransforms.accent1.redOff"
                                                    : greenModOnly
                                                        ? "$.design.theme.accentTransforms.accent1.greenMod"
                                                        : greenOffOnly
                                                            ? "$.design.theme.accentTransforms.accent1.greenOff"
                                                            : blueModOnly
                                                                ? "$.design.theme.accentTransforms.accent1.blueMod"
                                                                : blueOffOnly
                                                                    ? "$.design.theme.accentTransforms.accent1.blueOff"
                                                                    : hueModOnly
                                                                        ? "$.design.theme.accentTransforms.accent1.hueMod"
                                                                        : "$.design.theme.accentTransforms.accent1.hueOff";
            if (authoredTheme.AccentTransforms.Count != 1 ||
                transform.Role != "accent1" ||
                (!tintOnly && !shadeOnly && !lumModOnly && !lumOffOnly && !alphaModOnly && !alphaOffOnly && !satModOnly && !satOffOnly && !redModOnly && !redOffOnly && !greenModOnly && !greenOffOnly && !blueModOnly && !blueOffOnly && !hueModOnly && !hueOffOnly) ||
                !lumModOnly && transform.HasLuminanceModulationThousandth ||
                !lumOffOnly && transform.HasLuminanceOffsetThousandth ||
                !alphaModOnly && transform.HasAlphaModulationThousandth ||
                !alphaOffOnly && transform.HasAlphaOffsetThousandth ||
                !satModOnly && transform.HasSaturationModulationThousandth ||
                !satOffOnly && transform.HasSaturationOffsetThousandth ||
                !redModOnly && transform.HasRedModulationThousandth ||
                !redOffOnly && transform.HasRedOffsetThousandth ||
                !greenModOnly && transform.HasGreenModulationThousandth ||
                !greenOffOnly && transform.HasGreenOffsetThousandth ||
                !blueModOnly && transform.HasBlueModulationThousandth ||
                !blueOffOnly && transform.HasBlueOffsetThousandth ||
                !hueModOnly && transform.HasHueModulationThousandth ||
                !hueOffOnly && transform.HasHueOffsetAngleThousandth ||
                transform.HasGray || transform.HasComp || transform.HasInv || transform.HasGamma ||
                transform.HasInvGamma ||
                tintOnly && transform.TintThousandth > 100_000 ||
                shadeOnly && transform.ShadeThousandth > 100_000 ||
                lumModOnly && transform.LuminanceModulationThousandth > 100_000 ||
                lumOffOnly && (transform.LuminanceOffsetThousandth < -100_000 || transform.LuminanceOffsetThousandth > 100_000) ||
                alphaModOnly && transform.AlphaModulationThousandth > 100_000 ||
                alphaOffOnly && (transform.AlphaOffsetThousandth < -100_000 || transform.AlphaOffsetThousandth > 100_000) ||
                satModOnly && transform.SaturationModulationThousandth > 100_000 ||
                satOffOnly && (transform.SaturationOffsetThousandth < -100_000 || transform.SaturationOffsetThousandth > 100_000) ||
                redModOnly && transform.RedModulationThousandth > 100_000 ||
                redOffOnly && (transform.RedOffsetThousandth < -100_000 || transform.RedOffsetThousandth > 100_000) ||
                greenModOnly && transform.GreenModulationThousandth > 100_000 ||
                greenOffOnly && (transform.GreenOffsetThousandth < -100_000 || transform.GreenOffsetThousandth > 100_000) ||
                blueModOnly && transform.BlueModulationThousandth > 100_000 ||
                blueOffOnly && (transform.BlueOffsetThousandth < -100_000 || transform.BlueOffsetThousandth > 100_000) ||
                hueModOnly && transform.HueModulationThousandth > 100_000 ||
                hueOffOnly && (transform.HueOffsetAngleThousandth < -21_600_000 || transform.HueOffsetAngleThousandth > 21_600_000))
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export supports only one existing direct accent1 tint, shade, lumMod, lumOff, alphaMod, alphaOff, satMod, satOff, redMod, redOff, greenMod, greenOff, blueMod, blueOff, hueMod, or hueOff.",
                    transformPath);
            var colorScheme = theme.ThemeElements?.ColorScheme;
            var sourceValue = tintOnly
                ? (long?)TryReadSourceBoundThemeAccent1Tint(theme)
                : shadeOnly
                    ? (long?)TryReadSourceBoundThemeAccent1Shade(theme)
                    : lumModOnly
                        ? (long?)TryReadSourceBoundThemeAccent1LumMod(theme)
                        : lumOffOnly
                            ? (long?)TryReadSourceBoundThemeAccent1LumOff(theme)
                            : alphaModOnly
                            ? (long?)TryReadSourceBoundThemeAccent1AlphaMod(theme)
                                : alphaOffOnly
                            ? (long?)TryReadSourceBoundThemeAccent1AlphaOff(theme)
                                    : satModOnly
                                        ? (long?)TryReadSourceBoundThemeAccent1SatMod(theme)
                                        : satOffOnly
                                            ? (long?)TryReadSourceBoundThemeAccent1SatOff(theme)
                                            : redModOnly
                                                ? (long?)TryReadSourceBoundThemeAccent1RedMod(theme)
                                                : redOffOnly
                                                    ? (long?)TryReadSourceBoundThemeAccent1RedOff(theme)
                                                    : greenModOnly
                                                        ? (long?)TryReadSourceBoundThemeAccent1GreenMod(theme)
                                                        : greenOffOnly
                                                            ? (long?)TryReadSourceBoundThemeAccent1GreenOff(theme)
                                                            : blueModOnly
                                                                ? (long?)TryReadSourceBoundThemeAccent1BlueMod(theme)
                                                                    : blueOffOnly
                                                                        ? (long?)TryReadSourceBoundThemeAccent1BlueOff(theme)
                                                                        : hueModOnly
                                                                            ? (long?)TryReadSourceBoundThemeAccent1HueMod(theme)
                                                                            : (long?)TryReadSourceBoundThemeAccent1HueOff(theme);
            if (sourceValue is null)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    "Source-preserving PPTX export can edit only an existing direct accent1 tint, shade, lumMod, lumOff, alphaMod, alphaOff, satMod, satOff, redMod, redOff, greenMod, greenOff, blueMod, blueOff, hueMod, or hueOff.",
                    PartPath(themePart));
            var requested = tintOnly
                ? (long)transform.TintThousandth
                : shadeOnly
                    ? (long)transform.ShadeThousandth
                    : lumModOnly
                        ? (long)transform.LuminanceModulationThousandth
                        : lumOffOnly
                            ? (long)transform.LuminanceOffsetThousandth
                            : alphaModOnly
                                ? (long)transform.AlphaModulationThousandth
                                : alphaOffOnly
                                    ? (long)transform.AlphaOffsetThousandth
                                    : satModOnly
                                        ? (long)transform.SaturationModulationThousandth
                                        : satOffOnly
                                            ? (long)transform.SaturationOffsetThousandth
                                            : redModOnly
                                                ? (long)transform.RedModulationThousandth
                                                : redOffOnly
                                                    ? (long)transform.RedOffsetThousandth
                                                    : greenModOnly
                                                        ? (long)transform.GreenModulationThousandth
                                                        : greenOffOnly
                                                            ? (long)transform.GreenOffsetThousandth
                                                            : blueModOnly
                                                                ? (long)transform.BlueModulationThousandth
                                                            : blueOffOnly
                                                                ? (long)transform.BlueOffsetThousandth
                                                                : hueModOnly
                                                                    ? (long)transform.HueModulationThousandth
                                                                    : (long)transform.HueOffsetAngleThousandth;
            if (sourceValue.Value != requested)
            {
                var color = colorScheme?.Accent1Color?.GetFirstChild<A.RgbColorModelHex>() ??
                    throw new CodecException(
                        "unsupported_presentation_edit",
                        "Source-preserving PPTX export cannot create a missing direct accent1 color.",
                        PartPath(themePart));
                if (tintOnly)
                {
                    var tint = color.FirstChild as A.Tint ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 tint.",
                            PartPath(themePart));
                    tint.Val = checked((int)requested);
                }
                else if (shadeOnly)
                {
                    var shade = color.FirstChild as A.Shade ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 shade.",
                            PartPath(themePart));
                    shade.Val = checked((int)requested);
                }
                else if (lumModOnly)
                {
                    var lumMod = color.FirstChild as A.LuminanceModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 lumMod.",
                            PartPath(themePart));
                    lumMod.Val = checked((int)requested);
                }
                else if (lumOffOnly)
                {
                    var lumOff = color.FirstChild as A.LuminanceOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 lumOff.",
                            PartPath(themePart));
                    lumOff.Val = checked((int)requested);
                }
                else if (alphaModOnly)
                {
                    var alphaMod = color.FirstChild as A.AlphaModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 alphaMod.",
                            PartPath(themePart));
                    alphaMod.Val = checked((int)requested);
                }
                else if (alphaOffOnly)
                {
                    var alphaOff = color.FirstChild as A.AlphaOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 alphaOff.",
                            PartPath(themePart));
                    alphaOff.Val = checked((int)requested);
                }
                else if (satModOnly)
                {
                    var satMod = color.FirstChild as A.SaturationModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 satMod.",
                            PartPath(themePart));
                    satMod.Val = checked((int)requested);
                }
                else if (satOffOnly)
                {
                    var satOff = color.FirstChild as A.SaturationOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 satOff.",
                            PartPath(themePart));
                    satOff.Val = checked((int)requested);
                }
                else if (redModOnly)
                {
                    var redMod = color.FirstChild as A.RedModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 redMod.",
                            PartPath(themePart));
                    redMod.Val = checked((int)requested);
                }
                else if (redOffOnly)
                {
                    var redOff = color.FirstChild as A.RedOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 redOff.",
                        PartPath(themePart));
                    redOff.Val = checked((int)requested);
                }
                else if (greenModOnly)
                {
                    var greenMod = color.FirstChild as A.GreenModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 greenMod.",
                            PartPath(themePart));
                    greenMod.Val = checked((int)requested);
                }
                else if (greenOffOnly)
                {
                    var greenOff = color.FirstChild as A.GreenOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 greenOff.",
                            PartPath(themePart));
                    greenOff.Val = checked((int)requested);
                }
                else if (blueModOnly)
                {
                    var blueMod = color.FirstChild as A.BlueModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 blueMod.",
                            PartPath(themePart));
                    blueMod.Val = checked((int)requested);
                }
                else if (blueOffOnly)
                {
                    var blueOff = color.FirstChild as A.BlueOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 blueOff.",
                            PartPath(themePart));
                    blueOff.Val = checked((int)requested);
                }
                else if (hueModOnly)
                {
                    var hueMod = color.FirstChild as A.HueModulation ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 hueMod.",
                            PartPath(themePart));
                    hueMod.Val = checked((int)requested);
                }
                else
                {
                    var hueOff = color.FirstChild as A.HueOffset ??
                        throw new CodecException(
                            "unsupported_presentation_edit",
                            "Source-preserving PPTX export cannot create a missing direct accent1 hueOff.",
                            PartPath(themePart));
                    hueOff.Val = checked((int)requested);
                }
                changed = true;
            }
        }
        if (!changed) return;
        theme.Save();
        var partPath = PartPath(themePart);
        changedParts.Add(partPath);
        replacedOpaquePartHashes[partPath] = Hash(PartBytes(themePart));
    }

    private static string LayoutTypeName(P.SlideLayout source)
    {
        // `type` is a typed PresentationML attribute. Reading it through the
        // generic attribute accessor with an empty namespace is rejected by
        // Open XML SDK 3.x for typed elements, even though it is a valid
        // unqualified attribute in the serialized package. Use the generated
        // property so imported layouts from Office-authored templates remain
        // readable.
        var value = source.Type?.InnerText;
        return string.IsNullOrWhiteSpace(value) ? "custom" : value;
    }

    private static SlidePart[] ResolveSlideParts(PresentationPart presentationPart, IEnumerable<P.SlideId> slideIds) =>
        slideIds.Select(slideId => presentationPart.GetPartById(slideId.RelationshipId?.Value ?? string.Empty) as SlidePart ??
            throw new CodecException("missing_slide_part", "PPTX presentation contains an unresolved slide relationship.", "ppt/presentation.xml"))
        .ToArray();

    private static IReadOnlyDictionary<string, string> BuildCustomShowSlideIdMap(
        IEnumerable<(string RelationshipId, string PublicId)> entries)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var ambiguous = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (relationshipId, publicId) in entries)
        {
            if (relationshipId.Length == 0 || ambiguous.Contains(relationshipId)) continue;
            if (!result.TryAdd(relationshipId, publicId))
            {
                result.Remove(relationshipId);
                ambiguous.Add(relationshipId);
            }
        }
        return result;
    }

    private static IReadOnlyDictionary<uint, string> BuildSectionSlideIdMap(
        IEnumerable<(uint? NativeId, string PublicId)> entries)
    {
        var result = new Dictionary<uint, string>();
        var ambiguous = new HashSet<uint>();
        foreach (var (nativeId, publicId) in entries)
        {
            if (nativeId is not { } id || publicId.Length == 0 || ambiguous.Contains(id)) continue;
            if (!result.TryAdd(id, publicId))
            {
                result.Remove(id);
                ambiguous.Add(id);
            }
        }
        return result;
    }

    // Imported source topology keeps ordinary bindings and clone requests
    // separate. A clone is a new SlidePart with a verified origin, never a
    // second presentation reference to the same source part.
    private static PptxTargetSlideEntry[] BindSourcePreservingSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PresentationSlide> requested)
    {
        if (requested.Count == 0)
            throw new CodecException(
                "presentation_topology_changed",
                "Source-preserving PPTX export must retain at least one source slide.",
                "ppt/presentation.xml");
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var sourceSlides = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .ToArray();
        var targets = new PptxTargetSlideEntry[requested.Count];
        var seenSourceParts = new HashSet<SlidePart>();
        var seenCloneSourceParts = new HashSet<SlidePart>();
        for (var targetIndex = 0; targetIndex < requested.Count; targetIndex++)
        {
            var target = requested[targetIndex];
            var isClone = target.CloneSource is not null;
            if ((target.Source is null) == !isClone)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} must carry exactly one of source or clone_source.",
                    "ppt/presentation.xml");
            var binding = target.Source ?? target.CloneSource!;
            if (binding.SlideIndex >= sourceSlides.Length)
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} references source slide {binding.SlideIndex + 1}, which does not exist.",
                    "ppt/presentation.xml");
            var source = sourceSlides[binding.SlideIndex];
            var sourceRoot = source.Part.Slide ??
                throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
            if (!binding.PartPath.Equals(PartPath(source.Part), StringComparison.OrdinalIgnoreCase) ||
                !binding.RelationshipId.Equals(source.RelationshipId, StringComparison.Ordinal) ||
                !binding.SlideXmlSha256.Equals(HashElement(sourceRoot), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_slide_binding_mismatch",
                    $"Presentation slide {targetIndex + 1} does not match its hash-bound source slide.",
                    PartPath(source.Part));
            if (!isClone && !seenSourceParts.Add(source.Part))
                throw new CodecException(
                    "presentation_topology_changed",
                    "Source-preserving PPTX export cannot bind more than one ordinary target to a source SlidePart.",
                    "ppt/presentation.xml");
            if (isClone && !seenCloneSourceParts.Add(source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    "The bounded source-preserving PPTX clone profile permits only one pending clone per source SlidePart.",
                    PartPath(source.Part));
            targets[targetIndex] = new PptxTargetSlideEntry(targetIndex, target, source, isClone);
        }
        return targets;
    }

    private static void DeleteUnrequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<P.SlideId> sourceSlideIds,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        OpaqueOpcGraph opaque,
        ISet<string> changedParts,
        ISet<string> removedSourcePartPaths)
    {
        var sourceParts = ResolveSlideParts(presentationPart, sourceSlideIds);
        var retainedParts = targets.Where(target => !target.IsClone).Select(target => target.Source.Part).ToHashSet();
        var removed = sourceSlideIds
            .Select((slideId, index) => new PptxSourceSlideEntry(
                index,
                slideId,
                slideId.RelationshipId?.Value ?? string.Empty,
                sourceParts[index]))
            .Where(source => !retainedParts.Contains(source.Part))
            .ToArray();
        if (removed.Length == 0) return;

        var plans = removed
            .Select(source => (Source: source, Plan: PptxSlideDeletionCodec.Analyze(presentationPart, source, opaque)))
            .ToArray();
        foreach (var (source, plan) in plans)
        {
            if (!plan.Supported) throw UnsupportedSourceSlideDelete(source, plan.BlockedReason);
        }
        var transactionPlan = PptxSlideDeletionCodec.AnalyzeTransaction(presentationPart, removed, opaque);
        if (!transactionPlan.Supported)
            throw UnsupportedSourceSlideDelete(removed[0], transactionPlan.BlockedReason);

        foreach (var (source, _) in plans)
        {
            presentationPart.DeletePart(source.Part);
        }
        changedParts.UnionWith(transactionPlan.RemovedPackagePartPaths);
        removedSourcePartPaths.UnionWith(transactionPlan.RemovedPackagePartPaths);
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
    }

    private static CodecException UnsupportedSourceSlideDelete(PptxSourceSlideEntry source, string reason) =>
        new(
            "unsupported_presentation_slide_delete",
            $"Source-preserving PPTX deletion requires an exclusively owned OPC descendant closure; slide {source.Index + 1} cannot be deleted because {reason}.",
            PartPath(source.Part));

    private static void CloneRequestedSourceSlides(
        PresentationPart presentationPart,
        IReadOnlyList<PptxTargetSlideEntry> targets,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        PptxAssetCatalog assetCatalog,
        PptxCustomShowCatalog customShowCatalog,
        PptxNativeObjectCatalog nativeObjects,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths,
        ISet<string> clonedPackageEntryPaths,
        IDictionary<string, string> clonedPartSourcePaths)
    {
        var cloneTargets = targets.Where(target => target.IsClone).ToArray();
        if (cloneTargets.Length == 0) return;
        var retainedSlideParts = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var slideIdList = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var nextSlideId = slideIdList.Elements<P.SlideId>()
            .Select(slideId => slideId.Id?.Value ?? 255U)
            .DefaultIfEmpty(255U)
            .Max();

        foreach (var target in cloneTargets)
        {
            AssertSourceSlideRequestUnchanged(presentationPart, target, layoutIdByPartPath, slideIdByPartPath, assetCatalog, customShowCatalog, nativeObjects);
            var sourcePart = target.Source.Part;
            var omittedShapeTreeIndices = target.Target.ElementDeletions
                .Select(deletion => deletion.Source?.ShapeTreeIndex)
                .Where(index => index is not null)
                .Select(index => checked((int)index!.Value))
                .ToHashSet();
            var result = PptxSlideCloneCodec.Clone(presentationPart, target.Source, retainedSlideParts, omittedShapeTreeIndices);
            var clonePart = result.Part;
            // Validate the complete graph before applying any authorized
            // component projection. The clone codec proves that every
            // relationship and descendant is an exact source copy; the
            // bounded deletion pass below is the only permitted difference.
            PptxSlideCloneCodec.Validate(target.Source, clonePart, retainedSlideParts);
            changedParts.UnionWith(result.ChangedPackagePaths);
            clonedPackageEntryPaths.UnionWith(result.ChangedPackagePaths);
            addedPartPaths.UnionWith(result.AddedOpaquePartPaths);
            addedRelationshipIds.UnionWith(result.AddedOpaqueRelationshipKeys);
            foreach (var (clonePath, sourcePath) in result.CopiedPartSourcePaths)
                clonedPartSourcePaths.Add(clonePath, sourcePath);
            if (nextSlideId == uint.MaxValue)
                throw new CodecException("presentation_slide_id_exhausted", "PPTX cannot allocate another 32-bit slide identifier.", "ppt/presentation.xml");
            nextSlideId++;
            target.OutputPart = clonePart;
            target.OutputSlideId = new P.SlideId
            {
                Id = nextSlideId,
                RelationshipId = presentationPart.GetIdOfPart(clonePart),
            };
            ApplyCloneElementDeletions(
                target,
                clonePart,
                changedParts,
                addedRelationshipIds,
                addedPartPaths,
                clonedPackageEntryPaths,
                clonedPartSourcePaths);
        }
        changedParts.Add(PartPath(presentationPart));
        changedParts.Add(RelationshipPartPath(presentationPart));
        changedParts.Add("[Content_Types].xml");
        clonedPackageEntryPaths.Add("[Content_Types].xml");
    }

    private static void ApplyCloneElementDeletions(
        PptxTargetSlideEntry target,
        SlidePart clonePart,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths,
        ISet<string> clonedPackageEntryPaths,
        IDictionary<string, string> clonedPartSourcePaths)
    {
        if (target.Target.ElementDeletions.Count == 0) return;
        var sourceRoot = target.Source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {target.Source.Index + 1} has no slide root.", PartPath(target.Source.Part));
        var cloneRoot = clonePart.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation cloned slide {target.TargetIndex + 1} has no slide root.", PartPath(clonePart));
        var sourceElements = ShapeElements(sourceRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {target.Source.Index + 1} has no shape tree.", PartPath(target.Source.Part)));
        var cloneElements = ShapeElements(cloneRoot.CommonSlideData?.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation cloned slide {target.TargetIndex + 1} has no shape tree.", PartPath(clonePart)));
        if (sourceElements.Length != cloneElements.Length)
            throw PptxSlideCloneCodec.Unsupported(target.Source, "the cloned shape tree changed before the authorized component projection");

        var pending = new List<(int Index, OpenXmlElement Source, OpenXmlElement Clone, PptxElementDeletionPlan Plan)>();
        foreach (var deletion in target.Target.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", $"Presentation cloned slide {target.TargetIndex + 1} deletion {deletion.Id} is missing its source binding.", PartPath(target.Source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length)
                throw PptxSlideCloneCodec.Unsupported(target.Source, "an authorized component deletion identifies an invalid source shape-tree index");
            var sourceElement = sourceElements[sourceIndex];
            var cloneElement = cloneElements[sourceIndex];
            var sourcePlan = PptxElementDeletionCodec.Analyze(target.Source.Part, sourceElement, sourceElements, allowDuplicateNativeIds: true);
            var clonePlan = PptxElementDeletionCodec.Analyze(clonePart, cloneElement, cloneElements, allowDuplicateNativeIds: true);
            if (!sourcePlan.Supported || !clonePlan.Supported)
                throw new CodecException("unsupported_presentation_element_delete", $"Presentation cloned slide {target.TargetIndex + 1} component deletion {deletion.Id} is not supported by the source or clone deletion proof.", PartPath(clonePart));
            pending.Add((sourceIndex, sourceElement, cloneElement, clonePlan));
        }
        foreach (var deletion in pending.OrderByDescending(item => item.Index))
        {
            PptxElementDeletionCodec.Apply(clonePart, deletion.Clone, deletion.Plan);
            changedParts.Add(PartPath(clonePart));
            if (deletion.Plan.RelationshipIds.Count > 0)
            {
                changedParts.Add(RelationshipPartPath(clonePart));
                foreach (var relationshipId in deletion.Plan.RelationshipIds)
                    addedRelationshipIds.Remove($"{PartPath(clonePart)}\0{relationshipId}");
            }
            if (deletion.Plan.RemovedPackagePartPaths.Count > 0)
            {
                changedParts.UnionWith(deletion.Plan.RemovedPackagePartPaths);
                addedPartPaths.ExceptWith(deletion.Plan.RemovedPackagePartPaths);
                clonedPackageEntryPaths.ExceptWith(deletion.Plan.RemovedPackagePartPaths);
                foreach (var path in deletion.Plan.RemovedPackagePartPaths)
                    clonedPartSourcePaths.Remove(path);
                changedParts.Add("[Content_Types].xml");
            }
        }
        cloneRoot.Save();
    }

    // The first clone profile preserves the origin SlidePart byte-for-byte.
    // Allowing its removal in the same transaction would turn this leaf clone
    // into a topology-replacement primitive before the new graph has crossed
    // an export/reimport boundary, so reject it before any package mutation.
    private static void AssertCloneOriginsRetained(IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var retainedSources = targets
            .Where(target => !target.IsClone)
            .Select(target => target.Source.Part)
            .ToHashSet();
        foreach (var clone in targets.Where(target => target.IsClone))
        {
            if (!retainedSources.Contains(clone.Source.Part))
                throw new CodecException(
                    "unsupported_presentation_slide_clone",
                    $"Presentation clone {clone.TargetIndex + 1} cannot remove its origin slide in the same source-preserving export. Export and import the unchanged clone before changing source topology.",
                    PartPath(clone.Source.Part));
        }
    }

    private static void AssertSourceSlideRequestUnchanged(
        PresentationPart presentationPart,
        PptxTargetSlideEntry target,
        IReadOnlyDictionary<string, string> layoutIdByPartPath,
        IReadOnlyDictionary<string, string> slideIdByPartPath,
        PptxAssetCatalog assetCatalog,
        PptxCustomShowCatalog customShowCatalog,
        PptxNativeObjectCatalog nativeObjects)
    {
        var source = target.Source;
        var root = source.Part.Slide ??
            throw new CodecException("missing_slide_root", $"Presentation source slide {source.Index + 1} has no slide root.", PartPath(source.Part));
        var common = root.CommonSlideData ??
            throw new CodecException("missing_common_slide_data", $"Presentation source slide {source.Index + 1} has no common slide data.", PartPath(source.Part));
        var tree = common.ShapeTree ??
            throw new CodecException("missing_shape_tree", $"Presentation source slide {source.Index + 1} has no shape tree.", PartPath(source.Part));
        if (target.Target.Name != (common.Name?.Value ?? string.Empty))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source name");
        var layoutPart = source.Part.SlideLayoutPart ??
            throw PptxSlideCloneCodec.Unsupported(source, "it does not have a resolvable layout relationship");
        var expectedLayoutId = layoutIdByPartPath.GetValueOrDefault(PartPath(layoutPart));
        if (string.IsNullOrWhiteSpace(expectedLayoutId) || target.Target.LayoutId != expectedLayoutId)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source layout binding");
        var sourceVisibility = PptxSlideVisibilityCodec.Read(root);
        if (sourceVisibility.Editable != target.Target.HasHidden ||
            target.Target.HasHidden && sourceVisibility.Hidden != target.Target.Hidden)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes or invents its source visibility");
        var sourceBinding = target.Target.CloneSource ??
            throw new CodecException("missing_presentation_slide_clone_binding", $"Presentation clone {target.TargetIndex + 1} is missing clone_source.", PartPath(source.Part));
        var context = new PptxPartContext(source.Part, slideIdByPartPath, assets: assetCatalog, customShows: customShowCatalog);
        if (sourceBinding.LayoutRelationshipId != source.Part.GetIdOfPart(layoutPart) ||
            !sourceBinding.BackgroundSemanticSha256.Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common, context)), StringComparison.OrdinalIgnoreCase) ||
            !sourceBinding.TransitionSemanticSha256.Equals(PptxTransitionCodec.SemanticHash(PptxTransitionCodec.Read(root)), StringComparison.OrdinalIgnoreCase) ||
            sourceBinding.TransitionEditable != PptxTransitionCodec.Supports(root) ||
            sourceBinding.TransitionPresent != PptxTransitionCodec.HasTransition(root) ||
            sourceBinding.TransitionAddable != PptxTransitionCodec.CanAdd(root) ||
            sourceBinding.VisibilityEditable != sourceVisibility.Editable ||
            !sourceBinding.VisibilitySemanticSha256.Equals(sourceVisibility.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException("presentation_slide_clone_binding_mismatch", $"Presentation clone {target.TargetIndex + 1} does not match its source layout/background/transition binding.", PartPath(source.Part));
        if (!BackgroundSemanticHash(target.Target.Background).Equals(BackgroundSemanticHash(PptxBackgroundCodec.Read(common, context)), StringComparison.OrdinalIgnoreCase))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source background");
        if (!PptxTransitionCodec.SemanticHash(target.Target.Transition).Equals(PptxTransitionCodec.SemanticHash(PptxTransitionCodec.Read(root)), StringComparison.OrdinalIgnoreCase))
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone changes its source transition");
        if (!PptxSpeakerNotesCodec.Equivalent(target.Target.SpeakerNotes, PptxSpeakerNotesCodec.Read(source.Part)))
            throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} speaker notes are not unchanged source notes.", PartPath(source.Part));
        var legacyProfile = PptxLegacyCommentsCodec.Profile(presentationPart, source.Part, source.Index);
        if (legacyProfile.Supported && !PptxLegacyCommentsCodec.Equivalent(legacyProfile.Comments, target.Target.LegacyComments) ||
            !legacyProfile.Supported && target.Target.LegacyComments.Count > 0)
            throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} legacy comments are not unchanged source comments.", PartPath(source.Part));

        var sourceElements = ShapeElements(tree);
        var zOrderPlan = AnalyzeElementZOrder(sourceElements);
        if (sourceElements.Length != target.Target.Elements.Count + target.Target.ElementDeletions.Count)
            throw PptxSlideCloneCodec.Unsupported(source, "the requested clone does not account for every source element");
        var elementIdsByNativeId = NativeElementIds(sourceElements, $"presentation/slide/{source.Index + 1}");
        var requestedBySourceIndex = new Dictionary<int, PresentationElement>();
        var previousSourceIndex = -1;
        foreach (var requested in target.Target.Elements)
        {
            var binding = requested.Source ??
                throw new CodecException("missing_presentation_element_binding", $"Presentation clone {target.TargetIndex + 1} element {requested.Id} is missing its source binding.", PartPath(source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length || sourceIndex <= previousSourceIndex || !requestedBySourceIndex.TryAdd(sourceIndex, requested))
                throw PptxSlideCloneCodec.Unsupported(source, "retained clone elements do not preserve unique source shape-tree order");
            previousSourceIndex = sourceIndex;
        }
        var deletionsBySourceIndex = new Dictionary<int, PresentationElementDeletion>();
        foreach (var deletion in target.Target.ElementDeletions)
        {
            var binding = deletion.Source ??
                throw new CodecException("missing_presentation_element_deletion_binding", $"Presentation clone {target.TargetIndex + 1} deletion {deletion.Id} is missing its source binding.", PartPath(source.Part));
            var sourceIndex = checked((int)binding.ShapeTreeIndex);
            if (sourceIndex < 0 || sourceIndex >= sourceElements.Length || requestedBySourceIndex.ContainsKey(sourceIndex) || !deletionsBySourceIndex.TryAdd(sourceIndex, deletion))
                throw PptxSlideCloneCodec.Unsupported(source, "component deletions do not identify unique omitted source elements");
        }
        var retainedNativeIds = sourceElements
            .Select((element, index) => (element, index))
            .Where(item => !deletionsBySourceIndex.ContainsKey(item.index))
            .SelectMany(item => PptxElementDeletionCodec.NativeIds(item.element))
            .ToArray();
        if (retainedNativeIds.GroupBy(id => id).Any(group => group.Count() > 1))
            throw PptxSlideCloneCodec.Unsupported(source, "retained component elements contain duplicate native drawing IDs");
        var retainedNativeIdSet = retainedNativeIds.ToHashSet();
        for (var elementIndex = 0; elementIndex < sourceElements.Length; elementIndex++)
        {
            var original = ReadElement(sourceElements[elementIndex], source.Index, elementIndex, context, nativeObjects, elementIdsByNativeId);
            var strictDeletionPlan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[elementIndex], sourceElements);
            SetElementDeletionCapability(original, strictDeletionPlan);
            SetElementZOrderCapability(original, zOrderPlan);
            if (requestedBySourceIndex.TryGetValue(elementIndex, out var requested))
            {
                var binding = requested.Source!;
                AssertElementBinding(requested.Id, binding, sourceElements[elementIndex], original, strictDeletionPlan, zOrderPlan, target.TargetIndex, elementIndex, source.Part);
                if (HasParagraphTabStopRemovalIntent(requested) || HasParagraphRightMarginRemovalIntent(requested) || !SemanticHash(requested).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException("presentation_slide_clone_mismatch", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} is not an unchanged source element.", PartPath(source.Part));
                continue;
            }
            if (!deletionsBySourceIndex.TryGetValue(elementIndex, out var deletion))
                throw PptxSlideCloneCodec.Unsupported(source, "a source element is neither retained nor explicitly deleted");
            AssertElementBinding(deletion.Id, deletion.Source!, sourceElements[elementIndex], original, strictDeletionPlan, zOrderPlan, target.TargetIndex, elementIndex, source.Part);
            if (!deletion.Id.Equals(original.Id, StringComparison.Ordinal))
                throw new CodecException("presentation_element_deletion_binding_mismatch", $"Presentation clone {target.TargetIndex + 1} deletion {elementIndex + 1} changed its source element identity.", PartPath(source.Part));
            var projectionDeletionPlan = PptxElementDeletionCodec.Analyze(source.Part, sourceElements[elementIndex], sourceElements, allowDuplicateNativeIds: true);
            if (!projectionDeletionPlan.Supported)
                throw new CodecException("unsupported_presentation_element_delete", $"Presentation clone {target.TargetIndex + 1} element {elementIndex + 1} cannot be safely deleted: {projectionDeletionPlan.BlockedReason}.", PartPath(source.Part));
            if (PptxElementDeletionCodec.NativeIds(sourceElements[elementIndex]).Overlaps(retainedNativeIdSet))
                throw PptxSlideCloneCodec.Unsupported(source, "a deleted element shares a native drawing ID with a retained component element");
        }
    }

    private static bool ReorderSourceSlideIdList(PresentationPart presentationPart, IReadOnlyList<PptxTargetSlideEntry> targets)
    {
        var root = presentationPart.Presentation ??
            throw new CodecException("missing_presentation_root", "PPTX package has no Presentation root.", "ppt/presentation.xml");
        var list = root.SlideIdList ??
            throw new CodecException("missing_slide_id_list", "PPTX package has no slide ID list.", "ppt/presentation.xml");
        var sourceIds = list.Elements<P.SlideId>().ToArray();
        var requestedRelationshipIds = targets.Select(target => target.OutputSlideId.RelationshipId?.Value ?? string.Empty).ToArray();
        if (sourceIds.Select(slideId => slideId.RelationshipId?.Value ?? string.Empty).SequenceEqual(requestedRelationshipIds, StringComparer.Ordinal)) return false;

        // Keep any future extension child in its original relative position;
        // only replace the ordered p:sldId entries with exact source clones.
        var firstNonSlideId = list.ChildElements.FirstOrDefault(item => item is not P.SlideId);
        var reordered = targets.Select(target => (P.SlideId)target.OutputSlideId.CloneNode(true)).ToArray();
        foreach (var sourceId in sourceIds) sourceId.Remove();
        foreach (var slideId in reordered)
        {
            if (firstNonSlideId is null) list.Append(slideId);
            else list.InsertBefore(slideId, firstNonSlideId);
        }
        root.Save();
        return true;
    }

    private static string MasterTextStylesSemanticHash(PresentationMasterTextStyles? source)
    {
        var semantic = source?.Clone() ?? new PresentationMasterTextStyles();
        PptxMasterTextStylesCodec.NormalizeSemantics(semantic);
        return Hash(semantic.ToByteArray());
    }

    private static string BackgroundSemanticHash(PresentationBackground? source) =>
        Hash((source ?? new PresentationBackground()).ToByteArray());

    private static bool ApplyPlaceholders(
        P.ShapeTree shapeTree,
        string ownerId,
        IList<PresentationPlaceholder> requested,
        PptxPartContext partContext,
        string partPath)
    {
        var originals = PptxPlaceholderCodec.Read(shapeTree, ownerId, partContext);
        if (originals.Count != requested.Count)
            throw new CodecException(
                "presentation_placeholder_topology_changed",
                $"Source-preserving PPTX export requires {ownerId}'s original {originals.Count}-placeholder topology; the artifact contains {requested.Count} placeholders.",
                partPath);
        var changed = false;
        for (var index = 0; index < originals.Count; index++)
        {
            var original = originals[index];
            var target = requested[index];
            var sourceBinding = original.Source!;
            var binding = target.Source ?? throw new CodecException(
                "missing_presentation_placeholder_binding",
                $"Presentation placeholder {index + 1} under {ownerId} is missing its source binding.",
                partPath);
            var sourceShape = PptxPlaceholderCodec.BoundShape(shapeTree, original);
            if (sourceShape is null || target.Id != original.Id ||
                binding.ShapeTreeIndex != sourceBinding.ShapeTreeIndex ||
                !binding.ElementSha256.Equals(sourceBinding.ElementSha256, StringComparison.OrdinalIgnoreCase) ||
                !binding.SemanticSha256.Equals(sourceBinding.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
                binding.Editable != sourceBinding.Editable ||
                binding.DirectFramePresenceEditable != sourceBinding.DirectFramePresenceEditable ||
                binding.TextEditable != sourceBinding.TextEditable ||
                binding.AccessibilityEditable != sourceBinding.AccessibilityEditable ||
                !binding.ElementSha256.Equals(PptxPlaceholderCodec.ElementHash(sourceShape), StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_binding_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its hash-bound source element.",
                    partPath);
            if (!PptxPlaceholderCodec.SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "presentation_placeholder_source_semantics_mismatch",
                    $"Presentation placeholder {index + 1} under {ownerId} does not match its source semantic binding.",
                    partPath);
            if (PptxPlaceholderCodec.SemanticHash(target).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
            if (!binding.Editable)
                throw new CodecException(
                    "unsupported_presentation_edit",
                    $"Presentation placeholder {index + 1} under {ownerId} has no safely editable semantic component in this codec slice.",
                    partPath);
            PptxPlaceholderCodec.Apply(sourceShape, original, target, partContext);
            changed = true;
        }
        return changed;
    }

    private static IEnumerable<PresentationTextParagraph> MasterStyleParagraphs(PresentationMasterTextStyles? source) =>
        source is null
            ? []
            : source.TitleLevels.Concat(source.BodyLevels).Concat(source.OtherLevels);

    private static void TrackContextChanges(
        OpenXmlPart owner,
        PptxPartContext context,
        ISet<string> changedParts,
        ISet<string> addedRelationshipIds,
        ISet<string> addedPartPaths,
        ISet<string> removedRelationshipKeys,
        ISet<string> removedPartPaths)
    {
        if (context.RelationshipsChanged)
        {
            foreach (var key in context.AddedRelationshipKeys)
            {
                var separator = key.IndexOf('\0');
                if (separator <= 0 || separator == key.Length - 1) continue;
                addedRelationshipIds.Add(key);
                changedParts.Add(RelationshipPartPathFor(key[..separator]));
            }
            foreach (var key in context.RemovedRelationshipKeys)
            {
                var separator = key.IndexOf('\0');
                if (separator <= 0 || separator == key.Length - 1) continue;
                removedRelationshipKeys.Add(key);
                changedParts.Add(RelationshipPartPathFor(key[..separator]));
            }
        }
        foreach (var path in context.AddedPartPaths)
        {
            changedParts.Add(path);
            addedPartPaths.Add(path);
            changedParts.Add("[Content_Types].xml");
        }
        foreach (var path in context.RemovedPartPaths)
        {
            changedParts.Add(path);
            removedPartPaths.Add(path);
            changedParts.Add("[Content_Types].xml");
        }
    }

    private static string ShapeResidualHash(P.Shape source, PptxPartContext slideContext)
    {
        var shape = (P.Shape)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(shape);
        PptxNonVisualAccessibilityCodec.ScrubModeledContent(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);
        PptxHyperlinkCodec.ScrubElementAction(shape.NonVisualShapeProperties?.NonVisualDrawingProperties, slideContext);
        PptxHyperlinkCodec.ScrubElementHoverAction(shape.NonVisualShapeProperties?.NonVisualDrawingProperties, slideContext);
        if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties is { } nonVisual) nonVisual.Name = string.Empty;
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } drawingProperties) drawingProperties.TextBox = null;
        if (shape.ShapeProperties is { } properties)
        {
            var placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() is not null;
            if (properties.Transform2D is { } transform)
            {
                if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
                if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
                PptxShapeTransformCodec.Scrub(transform);
                // Adding the first owner-local a:xfrm is the bounded
                // materialization operation for an inherited slide
                // placeholder.  The frame itself is modeled separately, so
                // normalize transform presence in the residual hash while
                // keeping every non-placeholder topology change visible.
                if (placeholderShape) transform.Remove();
            }
            properties.GetFirstChild<A.CustomGeometry>()?.Remove();
            if (properties.GetFirstChild<A.PresetGeometry>() is { } geometry)
            {
                geometry.Preset = A.ShapeTypeValues.Rectangle;
                geometry.RemoveAllChildren();
                geometry.Append(new A.AdjustValueList());
            }
            else properties.InsertAfter(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }, properties.GetFirstChild<A.Transform2D>());
            properties.GetFirstChild<A.EffectList>()?.Remove();
            foreach (var fill in properties.ChildElements.Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill).ToArray()) fill.Remove();
            if (properties.GetFirstChild<A.Outline>() is { } outline)
                PptxLineStyleCodec.ScrubModeledContent(outline);
        }
        PptxTextCodec.ScrubModeledContent(shape.TextBody, slideContext);
        return HashElement(shape);
    }

    private static string ConnectorResidualHash(P.ConnectionShape source)
    {
        var connector = (P.ConnectionShape)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(connector);
        if (connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties is { } nonVisual)
        {
            PptxNonVisualAccessibilityCodec.ScrubModeledContent(nonVisual);
            nonVisual.Name = string.Empty;
        }
        if (connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties is { } drawingProperties)
            drawingProperties.RemoveAllChildren();
        connector.ShapeProperties?.RemoveAllChildren();
        return HashElement(connector);
    }

    private static string ChartFrameResidualHash(P.GraphicFrame source)
    {
        var chart = (P.GraphicFrame)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(chart);
        PptxChartCodec.ScrubFrame(chart);
        return HashElement(chart);
    }

    private static string PictureResidualHash(P.Picture source)
    {
        var picture = (P.Picture)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(picture);
        PptxPictureCodec.ScrubModeledContent(picture);
        return HashElement(picture);
    }

    private static string TableResidualHash(P.GraphicFrame source, PptxPartContext? context = null)
    {
        var table = (P.GraphicFrame)source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(table);
        PptxTableCodec.ScrubModeledContent(table, context);
        return HashElement(table);
    }

    private static string NativeObjectResidualHash(OpenXmlElement source)
    {
        var clone = source.CloneNode(true);
        PptxElementStateCodec.ScrubModeledContent(clone);
        if (clone is P.Picture mediaPicture &&
            PptxNativeObjectCatalog.IsMediaPicture(mediaPicture))
            PptxNonVisualAccessibilityCodec.ScrubResidualModeledContent(
                mediaPicture.NonVisualPictureProperties?.NonVisualDrawingProperties);
        if (clone.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault() is { } nonVisual)
            nonVisual.Name = string.Empty;
        if (clone is P.Picture picture && picture.ShapeProperties?.GetFirstChild<A.Transform2D>() is { } pictureTransform)
        {
            ScrubFrame(pictureTransform);
        }
        else if (clone is P.ConnectionShape connector && connector.ShapeProperties?.GetFirstChild<A.Transform2D>() is { } connectorTransform)
        {
            ScrubFrame(connectorTransform);
        }
        else if (clone is P.GraphicFrame graphicFrame && graphicFrame.Transform is { } transform)
        {
            ScrubFrame(transform);
            if (PptxNativeObjectCatalog.Classify(clone) == "oleObject" && graphicFrame.Descendants<A.Transform2D>().FirstOrDefault() is { } preview)
                ScrubFrame(preview);
        }
        else if (clone is P.GroupShape group && group.GetFirstChild<P.GroupShapeProperties>()?.GetFirstChild<A.TransformGroup>() is { } groupTransform)
        {
            ScrubFrame(groupTransform);
        }
        return HashElement(clone);
    }

    private static void ScrubFrame(P.Transform transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
        PptxFrameTransformCodec.Scrub(transform);
    }

    private static void ScrubFrame(A.Transform2D transform)
    {
        if (transform.Offset is { } offset) { offset.X = 0L; offset.Y = 0L; }
        if (transform.Extents is { } extents) { extents.Cx = 1L; extents.Cy = 1L; }
    }

    private static void ScrubFrame(A.TransformGroup transform)
    {
        transform.Offset!.X = 0L;
        transform.Offset.Y = 0L;
        transform.Extents!.Cx = 1L;
        transform.Extents.Cy = 1L;
        PptxFrameTransformCodec.Scrub(transform);
    }

    private static string MasterResidualHash(P.SlideMaster source, PptxPartContext partContext)
    {
        var master = (P.SlideMaster)source.CloneNode(true);
        PptxMasterTextStylesCodec.ScrubModeledContent(master, partContext);
        PptxBackgroundCodec.ScrubModeledContent(master.CommonSlideData, partContext);
        PptxPlaceholderCodec.ScrubModeledContent(master.CommonSlideData?.ShapeTree, partContext);
        return HashElement(master);
    }

    private static string LayoutResidualHash(P.SlideLayout source, PptxPartContext partContext)
    {
        var layout = (P.SlideLayout)source.CloneNode(true);
        PptxBackgroundCodec.ScrubModeledContent(layout.CommonSlideData, partContext);
        PptxPlaceholderCodec.ScrubModeledContent(layout.CommonSlideData?.ShapeTree, partContext);
        return HashElement(layout);
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => allowed.Contains(attribute.LocalName));
    }

    private static Dictionary<string, string> PackagePartHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.Where(entry => !entry.FullName.EndsWith('/')).ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var source = entry.Open();
                return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
            },
            StringComparer.OrdinalIgnoreCase);
    }

    private static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated PPTX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    private const int MaxOffice2021ValidationErrors = 256;

    private static void ValidateOffice2021(byte[] bytes)
    {
        var errors = Office2021ValidationErrors(bytes);
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Take(8).Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Generated PPTX is not valid Office 2021 Open XML: {detail}");
    }

    // An imported package may use valid-in-host extension markup that the
    // Open XML SDK's Office 2021 validator does not recognize. Source-bound
    // export may retain such diagnostics, but must never introduce a new one.
    // This comparison is intentionally unavailable to source-free export.
    private static int ValidateOffice2021AgainstSource(
        byte[] sourceBytes,
        byte[] outputBytes,
        IReadOnlyDictionary<string, string>? clonedPartSourcePaths = null)
    {
        var sourceErrors = Office2021ValidationErrors(sourceBytes);
        var sourceSignatures = sourceErrors
            .Select(error => error.Signature)
            .ToHashSet(StringComparer.Ordinal);
        var exactCopiedPartSources = ExactCopiedPartSources(sourceBytes, outputBytes, clonedPartSourcePaths);
        var introduced = Office2021ValidationErrors(outputBytes)
            .Where(error =>
                !sourceSignatures.Contains(error.Signature) &&
                (!exactCopiedPartSources.TryGetValue(error.Part, out var sourcePart) ||
                 !sourceSignatures.Contains(error.SignatureForPart(sourcePart))))
            .Take(8)
            .ToArray();
        if (introduced.Length == 0) return sourceErrors.Length;
        var detail = string.Join("; ", introduced.Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Source-preserving PPTX export introduced Office 2021 Open XML validation error(s): {detail}");
    }

    private sealed record Office2021ValidationError(
        string Part,
        string Path,
        string Id,
        string Description)
    {
        internal string Signature => SignatureForPart(Part);
        internal string Detail => $"{(Path.Length == 0 ? Part : Path)}: {Description}";
        internal string SignatureForPart(string part) => string.Join("\u001f", part, Path, Id, Description);
    }

    private static Office2021ValidationError[] Office2021ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021)
            .Validate(package)
            .Where(error => !IsCanonicalMathValidatorFalsePositive(error))
            .Take(MaxOffice2021ValidationErrors + 1)
            .Select(error =>
            {
                var part = error.Part?.Uri.ToString().TrimStart('/') ?? "package";
                var path = error.Path?.XPath ?? string.Empty;
                return new Office2021ValidationError(part, path, error.Id ?? string.Empty, error.Description ?? string.Empty);
            })
            .ToArray();
        if (errors.Length <= MaxOffice2021ValidationErrors) return errors;
        throw new CodecException(
            "openxml_validation_budget_exceeded",
            $"PPTX validation produced more than {MaxOffice2021ValidationErrors} errors; refusing an unbounded validation result.");
    }

    private static bool IsCanonicalMathValidatorFalsePositive(ValidationErrorInfo error)
    {
        if (error.Part is not SlidePart ||
            error.Path?.XPath?.Contains("/a14:m[", StringComparison.Ordinal) != true ||
            error.Description?.Contains("leaf element and cannot contain children", StringComparison.Ordinal) != true)
            return false;
        try
        {
            using var partStream = error.Part.GetStream(FileMode.Open, FileAccess.Read);
            var document = XDocument.Load(partStream, LoadOptions.PreserveWhitespace);
            XNamespace a14 = "http://schemas.microsoft.com/office/drawing/2010/main";
            var formulas = document.Descendants(a14 + "m").ToArray();
            return formulas.Length > 0 && formulas.All(PptxMathCodec.IsCanonical);
        }
        catch (Exception exception) when (exception is IOException or InvalidOperationException or System.Xml.XmlException)
        {
            return false;
        }
    }

    private static IReadOnlyDictionary<string, string> ExactCopiedPartSources(
        byte[] sourceBytes,
        byte[] outputBytes,
        IReadOnlyDictionary<string, string>? clonedPartSourcePaths)
    {
        if (clonedPartSourcePaths is null || clonedPartSourcePaths.Count == 0)
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sourceHashes = PackagePartHashes(sourceBytes);
        var outputHashes = PackagePartHashes(outputBytes);
        return clonedPartSourcePaths
            .Where(pair =>
                sourceHashes.TryGetValue(pair.Value, out var sourceHash) &&
                outputHashes.TryGetValue(pair.Key, out var outputHash) &&
                sourceHash.Equals(outputHash, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
    }

}
