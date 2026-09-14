using System.Buffers;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal sealed record PpjProjectionResult(
    PresentationProgramResult Program,
    IReadOnlyList<Diagnostic> Diagnostics,
    ArtifactEnvelope? SourceArtifact,
    IReadOnlyDictionary<string, PpjNativeLeafBinding> NativeLeafBindings,
    PpjValidationResult? Validation) : IDisposable
{
    public void Dispose() => Validation?.Dispose();
    internal IReadOnlyList<PptxNativeBinding> NativeBindings { get; init; } = [];
}

/// <summary>
/// Projects a validated PPTX package into the bounded public PPJ language.
/// Native package locators and raw OOXML deliberately stay behind opaque,
/// source-bound handles; the program exposes only semantic state and
/// hash-bound capabilities that the existing source-preserving writer can
/// independently prove again during compilation.
/// </summary>
internal static partial class PpjPresentationProjector
{
    private const double EmuPerPoint = 12_700d;

    internal static PpjProjectionResult Project(
        byte[] sourceBytes,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits,
        bool retainSourceAssetData = true,
        string? verifiedSourceSha256 = null,
        bool includeNativeBindings = false) => Project(
            new PptxPackageSource(sourceBytes),
            request,
            limits,
            retainSourceAssetData,
            verifiedSourceSha256,
            includeNativeBindings);

    internal static PpjProjectionResult Project(
        PptxPackageSource source,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits,
        bool retainSourceAssetData = true,
        string? verifiedSourceSha256 = null,
        bool includeNativeBindings = false)
    {
        if (PpjEmbeddedProgramCodec.TryRecover(source, request, limits) is { } recovered)
            return new(
                recovered.Program,
                recovered.Diagnostics,
                null,
                new Dictionary<string, PpjNativeLeafBinding>(StringComparer.Ordinal),
                null);

        var imported = PptxCodec.Import(
            source,
            limits,
            retainSourceAssetData,
            verifiedSourceSha256,
            includeNativeBindings);
        var envelope = imported.Artifact;
        var presentation = envelope.Presentation ??
            throw new CodecException("ppj.projection.presentation", "The imported package did not produce a Presentation artifact.", "$");
        var sourceSha256 = envelope.Source?.PackageSha256;
        if (string.IsNullOrEmpty(sourceSha256))
            sourceSha256 = source.Sha256();
        var revision = $"pptx-{sourceSha256[..16]}";
        var sourceUri = string.IsNullOrWhiteSpace(request.SourceUri)
            ? $"deck.assets/source/{sourceSha256}.pptx"
            : request.SourceUri;
        var assetRoot = string.IsNullOrWhiteSpace(request.AssetRootUri)
            ? "deck.assets/media"
            : request.AssetRootUri.TrimEnd('/');

        var context = new ProjectionContext(sourceSha256, revision, assetRoot, envelope.Assets, envelope.OpaqueOpc, source);
        RegisterIds(presentation, context);

        var pages = new JsonArray();
        foreach (var slide in presentation.Slides)
            pages.Add(ProjectPage(slide, presentation, context));

        foreach (var nativeAsset in context.NativeSourceAssets)
            if (!envelope.Assets.Any(asset => asset.Id.Equals(nativeAsset.Id, StringComparison.Ordinal)))
                envelope.Assets.Add(nativeAsset.Clone());

        var assets = context.ProgramAssets;
        var sections = ProjectSections(presentation, context);
        var customShows = ProjectCustomShows(presentation, context);
        var comments = ProjectComments(presentation, context);
        var nodeMap = context.BuildNodeMap();
        var nodeMapBytes = CanonicalBytes(nodeMap);
        var projectionPayload = new JsonObject
        {
            ["canvas"] = FrameDimensions(presentation),
            ["assets"] = assets,
            ["pages"] = pages,
            ["sections"] = sections,
            ["customShows"] = customShows,
            ["comments"] = comments,
        };
        var projectionSha256 = Sha256(CanonicalBytes(projectionPayload));
        // JsonNode has single-parent ownership. The payload exists only to
        // bind the source-derived semantic graph, so release its children and
        // reuse those exact nodes in the public program instead of cloning the
        // full projection.
        projectionPayload.Clear();

        var root = new JsonObject
        {
            ["schema"] = StringNode("office-kit/ppj/v1"),
            ["meta"] = new JsonObject
            {
                ["id"] = StringNode(StableDocumentId(presentation.Id, sourceSha256)),
                ["title"] = StringNode(string.IsNullOrWhiteSpace(presentation.Name) ? "Imported presentation" : presentation.Name),
                ["language"] = StringNode("und"),
                ["version"] = JsonValue.Create(1),
                ["description"] = StringNode("Source-derived PPJ projection. Unmodeled native content remains in the hash-bound PPTX source package."),
            },
            ["intent"] = ImportedIntent(),
            ["design"] = ImportedDesign(presentation, context),
            ["assets"] = assets,
            ["source"] = new JsonObject
            {
                ["kind"] = StringNode("pptx"),
                ["uri"] = StringNode(sourceUri),
                ["sha256"] = StringNode(sourceSha256),
                ["revision"] = StringNode(revision),
                ["projection"] = new JsonObject
                {
                    ["version"] = JsonValue.Create(1),
                    ["sha256"] = StringNode(projectionSha256),
                    ["nodeMapSha256"] = StringNode(Sha256(nodeMapBytes)),
                    ["visibleObjectCount"] = JsonValue.Create(context.VisibleObjectCount),
                },
            },
            ["pages"] = pages,
        };
        if (sections.Count > 0) root["sections"] = sections;
        if (customShows.Count > 0) root["customShows"] = customShows;
        if (comments.Count > 0) root["comments"] = comments;

        var candidateBytes = CanonicalBytes(root);
        root.Clear();
        pages.Clear();
        sections.Clear();
        customShows.Clear();
        comments.Clear();
        nodeMap.Clear();
        context.ReleaseProjectionJson();
        var validation = PpjProgramValidator.Validate(candidateBytes);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            validation.Dispose();
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        var result = new PresentationProgramResult
        {
            ProgramJson = UnsafeByteOperations.UnsafeWrap(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? UnsafeByteOperations.UnsafeWrap(nodeMapBytes) : ByteString.Empty,
            SourceSha256 = sourceSha256,
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        result.Assets.Add(context.ResultAssets);
        return new(result, imported.Diagnostics, envelope, context.NativeLeafBindings, validation)
        {
            NativeBindings = includeNativeBindings ? imported.NativeBindings
                .Where(binding => context.TryElementId(context.PageId(binding.PageId), binding.ElementId, out _))
                .Select(binding => binding with
                {
                    PageId = context.PageId(binding.PageId),
                    ElementId = context.ElementId(context.PageId(binding.PageId), binding.ElementId),
                }).ToArray() : [],
        };
    }

    private static JsonObject ImportedIntent() => new()
    {
        ["brief"] = new JsonObject
        {
            ["primaryJob"] = StringNode("handoff"),
            ["expectedOutcome"] = StringNode("Continue editing this presentation while preserving source-owned native content."),
            ["evidenceBoundary"] = StringNode("The source presentation does not declare its audience, factual authority, or intended outcome."),
        },
        ["audience"] = new JsonObject
        {
            ["description"] = StringNode("Imported presentation audience was not declared."),
        },
        ["narrative"] = new JsonObject
        {
            ["thesis"] = StringNode("Preserve and continue the imported presentation without inventing its original intent."),
        },
        ["editorial"] = new JsonObject
        {
            ["tone"] = new JsonArray { StringNode("source-derived") },
            ["avoid"] = new JsonArray { StringNode("Inventing missing facts or silently rewriting source-owned content") },
        },
        ["delivery"] = new JsonObject
        {
            ["mode"] = StringNode("hybrid"),
            ["mediumFit"] = StringNode("acceptable"),
            ["mediumFitNote"] = StringNode("Delivery intent was not declared in the imported file."),
        },
    };

    private static JsonObject ImportedDesign(PresentationArtifact presentation, ProjectionContext context)
    {
        var theme = new JsonObject
        {
            ["name"] = StringNode(presentation.AuthoredTheme?.HasName == true
                ? presentation.AuthoredTheme.Name
                : "Source-owned presentation theme"),
            // Imported run and bullet styles may retain a direct theme token
            // even though the source theme graph is opaque to the bounded
            // writer. Keep every standard token addressable with a neutral
            // fallback so the PPJ projection stays valid without claiming
            // that these fallback RGB values replace the source theme.
            ["colors"] = ImportedThemeColors(),
        };
        var themeCapabilities = new List<CapabilitySpec>();
        if (presentation.AuthoredTheme?.HasName == true)
            themeCapabilities.Add(new("setThemeName", ["name"]));
        if (presentation.AuthoredTheme is
        {
            HasDark1Rgb: true,
            HasLight1Rgb: true,
            HasDark2Rgb: true,
            HasLight2Rgb: true,
            HasHyperlinkRgb: true,
            HasFollowedHyperlinkRgb: true,
        } colorRoleTheme)
        {
            theme["colorRoles"] = new JsonObject
            {
                ["dark1"] = StringNode("#" + colorRoleTheme.Dark1Rgb),
                ["light1"] = StringNode("#" + colorRoleTheme.Light1Rgb),
                ["dark2"] = StringNode("#" + colorRoleTheme.Dark2Rgb),
                ["light2"] = StringNode("#" + colorRoleTheme.Light2Rgb),
                ["hyperlink"] = StringNode("#" + colorRoleTheme.HyperlinkRgb),
                ["followedHyperlink"] = StringNode("#" + colorRoleTheme.FollowedHyperlinkRgb),
            };
            themeCapabilities.Add(new("setThemeColorRoleDark1", ["colorRoles.dark1"]));
            themeCapabilities.Add(new("setThemeColorRoleLight1", ["colorRoles.light1"]));
            themeCapabilities.Add(new("setThemeColorRoleDark2", ["colorRoles.dark2"]));
            themeCapabilities.Add(new("setThemeColorRoleLight2", ["colorRoles.light2"]));
            themeCapabilities.Add(new("setThemeColorRoleHyperlink", ["colorRoles.hyperlink"]));
            themeCapabilities.Add(new("setThemeColorRoleFollowedHyperlink", ["colorRoles.followedHyperlink"]));
        }
        if (presentation.AuthoredTheme is { AccentRgb.Count: 6 } authoredTheme)
        {
            var accentColors = new JsonObject();
            var roles = new[] { "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" };
            for (var index = 0; index < roles.Length; index++)
                accentColors[roles[index]] = StringNode("#" + authoredTheme.AccentRgb[index]);
            theme["accentColors"] = accentColors;
            themeCapabilities.Add(new("setThemeAccent1Color", ["accentColors.accent1"]));
            themeCapabilities.Add(new("setThemeAccent2Color", ["accentColors.accent2"]));
            themeCapabilities.Add(new("setThemeAccent3Color", ["accentColors.accent3"]));
            themeCapabilities.Add(new("setThemeAccent4Color", ["accentColors.accent4"]));
            themeCapabilities.Add(new("setThemeAccent5Color", ["accentColors.accent5"]));
            themeCapabilities.Add(new("setThemeAccent6Color", ["accentColors.accent6"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } transformTheme &&
            transformTheme.AccentTransforms[0] is { Role: "accent1", HasTintThousandth: true } accent1Tint &&
            !accent1Tint.HasShadeThousandth &&
            !accent1Tint.HasLuminanceModulationThousandth &&
            !accent1Tint.HasLuminanceOffsetThousandth &&
            !accent1Tint.HasAlphaModulationThousandth &&
            !accent1Tint.HasAlphaOffsetThousandth &&
            !accent1Tint.HasSaturationModulationThousandth &&
            !accent1Tint.HasSaturationOffsetThousandth &&
            !accent1Tint.HasRedModulationThousandth &&
            !accent1Tint.HasRedOffsetThousandth &&
            !accent1Tint.HasGreenModulationThousandth &&
            !accent1Tint.HasGreenOffsetThousandth &&
            !accent1Tint.HasBlueModulationThousandth &&
            !accent1Tint.HasBlueOffsetThousandth &&
            !accent1Tint.HasHueModulationThousandth &&
            !accent1Tint.HasHueOffsetAngleThousandth &&
            !accent1Tint.HasGray && !accent1Tint.HasComp && !accent1Tint.HasInv &&
            !accent1Tint.HasGamma && !accent1Tint.HasInvGamma &&
            accent1Tint.TintThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent1Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1Tint", ["accentTransforms.accent1.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } shadeTheme &&
            shadeTheme.AccentTransforms[0] is { Role: "accent1", HasShadeThousandth: true } accent1Shade &&
            !accent1Shade.HasTintThousandth &&
            !accent1Shade.HasLuminanceModulationThousandth &&
            !accent1Shade.HasLuminanceOffsetThousandth &&
            !accent1Shade.HasAlphaModulationThousandth &&
            !accent1Shade.HasAlphaOffsetThousandth &&
            !accent1Shade.HasSaturationModulationThousandth &&
            !accent1Shade.HasSaturationOffsetThousandth &&
            !accent1Shade.HasRedModulationThousandth &&
            !accent1Shade.HasRedOffsetThousandth &&
            !accent1Shade.HasGreenModulationThousandth &&
            !accent1Shade.HasGreenOffsetThousandth &&
            !accent1Shade.HasBlueModulationThousandth &&
            !accent1Shade.HasBlueOffsetThousandth &&
            !accent1Shade.HasHueModulationThousandth &&
            !accent1Shade.HasHueOffsetAngleThousandth &&
            !accent1Shade.HasGray && !accent1Shade.HasComp && !accent1Shade.HasInv &&
            !accent1Shade.HasGamma && !accent1Shade.HasInvGamma &&
            accent1Shade.ShadeThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent1Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1Shade", ["accentTransforms.accent1.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } lumModTheme &&
            lumModTheme.AccentTransforms[0] is { Role: "accent1", HasLuminanceModulationThousandth: true } accent1LumMod &&
            !accent1LumMod.HasTintThousandth &&
            !accent1LumMod.HasShadeThousandth &&
            !accent1LumMod.HasLuminanceOffsetThousandth &&
            !accent1LumMod.HasAlphaModulationThousandth &&
            !accent1LumMod.HasAlphaOffsetThousandth &&
            !accent1LumMod.HasSaturationModulationThousandth &&
            !accent1LumMod.HasSaturationOffsetThousandth &&
            !accent1LumMod.HasRedModulationThousandth &&
            !accent1LumMod.HasRedOffsetThousandth &&
            !accent1LumMod.HasGreenModulationThousandth &&
            !accent1LumMod.HasGreenOffsetThousandth &&
            !accent1LumMod.HasBlueModulationThousandth &&
            !accent1LumMod.HasBlueOffsetThousandth &&
            !accent1LumMod.HasHueModulationThousandth &&
            !accent1LumMod.HasHueOffsetAngleThousandth &&
            !accent1LumMod.HasGray && !accent1LumMod.HasComp && !accent1LumMod.HasInv &&
            !accent1LumMod.HasGamma && !accent1LumMod.HasInvGamma &&
            accent1LumMod.LuminanceModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent1LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1LumMod", ["accentTransforms.accent1.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } lumOffTheme &&
            lumOffTheme.AccentTransforms[0] is { Role: "accent1", HasLuminanceOffsetThousandth: true } accent1LumOff &&
            !accent1LumOff.HasTintThousandth &&
            !accent1LumOff.HasShadeThousandth &&
            !accent1LumOff.HasLuminanceModulationThousandth &&
            !accent1LumOff.HasAlphaModulationThousandth &&
            !accent1LumOff.HasAlphaOffsetThousandth &&
            !accent1LumOff.HasSaturationModulationThousandth &&
            !accent1LumOff.HasSaturationOffsetThousandth &&
            !accent1LumOff.HasRedModulationThousandth &&
            !accent1LumOff.HasRedOffsetThousandth &&
            !accent1LumOff.HasGreenModulationThousandth &&
            !accent1LumOff.HasGreenOffsetThousandth &&
            !accent1LumOff.HasBlueModulationThousandth &&
            !accent1LumOff.HasBlueOffsetThousandth &&
            !accent1LumOff.HasHueModulationThousandth &&
            !accent1LumOff.HasHueOffsetAngleThousandth &&
            !accent1LumOff.HasGray && !accent1LumOff.HasComp && !accent1LumOff.HasInv &&
            !accent1LumOff.HasGamma && !accent1LumOff.HasInvGamma &&
            accent1LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent1LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent1LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1LumOff", ["accentTransforms.accent1.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } alphaModTheme &&
            alphaModTheme.AccentTransforms[0] is { Role: "accent1", HasAlphaModulationThousandth: true } accent1AlphaMod &&
            !accent1AlphaMod.HasTintThousandth &&
            !accent1AlphaMod.HasShadeThousandth &&
            !accent1AlphaMod.HasLuminanceModulationThousandth &&
            !accent1AlphaMod.HasLuminanceOffsetThousandth &&
            !accent1AlphaMod.HasAlphaOffsetThousandth &&
            !accent1AlphaMod.HasSaturationModulationThousandth &&
            !accent1AlphaMod.HasSaturationOffsetThousandth &&
            !accent1AlphaMod.HasRedModulationThousandth &&
            !accent1AlphaMod.HasRedOffsetThousandth &&
            !accent1AlphaMod.HasGreenModulationThousandth &&
            !accent1AlphaMod.HasGreenOffsetThousandth &&
            !accent1AlphaMod.HasBlueModulationThousandth &&
            !accent1AlphaMod.HasBlueOffsetThousandth &&
            !accent1AlphaMod.HasHueModulationThousandth &&
            !accent1AlphaMod.HasHueOffsetAngleThousandth &&
            !accent1AlphaMod.HasGray && !accent1AlphaMod.HasComp && !accent1AlphaMod.HasInv &&
            !accent1AlphaMod.HasGamma && !accent1AlphaMod.HasInvGamma &&
            accent1AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent1AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1AlphaMod", ["accentTransforms.accent1.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } alphaOffTheme &&
            alphaOffTheme.AccentTransforms[0] is { Role: "accent1", HasAlphaOffsetThousandth: true } accent1AlphaOff &&
            !accent1AlphaOff.HasTintThousandth &&
            !accent1AlphaOff.HasShadeThousandth &&
            !accent1AlphaOff.HasLuminanceModulationThousandth &&
            !accent1AlphaOff.HasLuminanceOffsetThousandth &&
            !accent1AlphaOff.HasAlphaModulationThousandth &&
            !accent1AlphaOff.HasSaturationModulationThousandth &&
            !accent1AlphaOff.HasSaturationOffsetThousandth &&
            !accent1AlphaOff.HasRedModulationThousandth &&
            !accent1AlphaOff.HasRedOffsetThousandth &&
            !accent1AlphaOff.HasGreenModulationThousandth &&
            !accent1AlphaOff.HasGreenOffsetThousandth &&
            !accent1AlphaOff.HasBlueModulationThousandth &&
            !accent1AlphaOff.HasBlueOffsetThousandth &&
            !accent1AlphaOff.HasHueModulationThousandth &&
            !accent1AlphaOff.HasHueOffsetAngleThousandth &&
            !accent1AlphaOff.HasGray && !accent1AlphaOff.HasComp && !accent1AlphaOff.HasInv &&
            !accent1AlphaOff.HasGamma && !accent1AlphaOff.HasInvGamma &&
            accent1AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent1AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent1AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1AlphaOff", ["accentTransforms.accent1.alphaOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } satModTheme &&
            satModTheme.AccentTransforms[0] is { Role: "accent1", HasSaturationModulationThousandth: true } accent1SatMod &&
            !accent1SatMod.HasTintThousandth &&
            !accent1SatMod.HasShadeThousandth &&
            !accent1SatMod.HasLuminanceModulationThousandth &&
            !accent1SatMod.HasLuminanceOffsetThousandth &&
            !accent1SatMod.HasAlphaModulationThousandth &&
            !accent1SatMod.HasAlphaOffsetThousandth &&
            !accent1SatMod.HasSaturationOffsetThousandth &&
            !accent1SatMod.HasRedModulationThousandth &&
            !accent1SatMod.HasRedOffsetThousandth &&
            !accent1SatMod.HasGreenModulationThousandth &&
            !accent1SatMod.HasGreenOffsetThousandth &&
            !accent1SatMod.HasBlueModulationThousandth &&
            !accent1SatMod.HasBlueOffsetThousandth &&
            !accent1SatMod.HasHueModulationThousandth &&
            !accent1SatMod.HasHueOffsetAngleThousandth &&
            !accent1SatMod.HasGray && !accent1SatMod.HasComp && !accent1SatMod.HasInv &&
            !accent1SatMod.HasGamma && !accent1SatMod.HasInvGamma &&
            accent1SatMod.SaturationModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent1SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1SatMod", ["accentTransforms.accent1.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } satOffTheme &&
            satOffTheme.AccentTransforms[0] is { Role: "accent1", HasSaturationOffsetThousandth: true } accent1SatOff &&
            !accent1SatOff.HasTintThousandth &&
            !accent1SatOff.HasShadeThousandth &&
            !accent1SatOff.HasLuminanceModulationThousandth &&
            !accent1SatOff.HasLuminanceOffsetThousandth &&
            !accent1SatOff.HasAlphaModulationThousandth &&
            !accent1SatOff.HasAlphaOffsetThousandth &&
            !accent1SatOff.HasSaturationModulationThousandth &&
            !accent1SatOff.HasRedModulationThousandth &&
            !accent1SatOff.HasRedOffsetThousandth &&
            !accent1SatOff.HasGreenModulationThousandth &&
            !accent1SatOff.HasGreenOffsetThousandth &&
            !accent1SatOff.HasBlueModulationThousandth &&
            !accent1SatOff.HasBlueOffsetThousandth &&
            !accent1SatOff.HasHueModulationThousandth &&
            !accent1SatOff.HasHueOffsetAngleThousandth &&
            !accent1SatOff.HasGray && !accent1SatOff.HasComp && !accent1SatOff.HasInv &&
            !accent1SatOff.HasGamma && !accent1SatOff.HasInvGamma &&
            accent1SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent1SatOff.SaturationOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent1SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1SatOff", ["accentTransforms.accent1.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } redModTheme &&
            redModTheme.AccentTransforms[0] is { Role: "accent1", HasRedModulationThousandth: true } accent1RedMod &&
            !accent1RedMod.HasTintThousandth &&
            !accent1RedMod.HasShadeThousandth &&
            !accent1RedMod.HasLuminanceModulationThousandth &&
            !accent1RedMod.HasLuminanceOffsetThousandth &&
            !accent1RedMod.HasAlphaModulationThousandth &&
            !accent1RedMod.HasAlphaOffsetThousandth &&
            !accent1RedMod.HasSaturationModulationThousandth &&
            !accent1RedMod.HasSaturationOffsetThousandth &&
            !accent1RedMod.HasRedOffsetThousandth &&
            !accent1RedMod.HasGreenModulationThousandth &&
            !accent1RedMod.HasGreenOffsetThousandth &&
            !accent1RedMod.HasBlueModulationThousandth &&
            !accent1RedMod.HasBlueOffsetThousandth &&
            !accent1RedMod.HasHueModulationThousandth &&
            !accent1RedMod.HasHueOffsetAngleThousandth &&
            !accent1RedMod.HasGray && !accent1RedMod.HasComp && !accent1RedMod.HasInv &&
            !accent1RedMod.HasGamma && !accent1RedMod.HasInvGamma &&
            accent1RedMod.RedModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent1RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1RedMod", ["accentTransforms.accent1.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } redOffTheme &&
            redOffTheme.AccentTransforms[0] is { Role: "accent1", HasRedOffsetThousandth: true } accent1RedOff &&
            !accent1RedOff.HasTintThousandth &&
            !accent1RedOff.HasShadeThousandth &&
            !accent1RedOff.HasLuminanceModulationThousandth &&
            !accent1RedOff.HasLuminanceOffsetThousandth &&
            !accent1RedOff.HasAlphaModulationThousandth &&
            !accent1RedOff.HasAlphaOffsetThousandth &&
            !accent1RedOff.HasSaturationModulationThousandth &&
            !accent1RedOff.HasSaturationOffsetThousandth &&
            !accent1RedOff.HasRedModulationThousandth &&
            !accent1RedOff.HasGreenModulationThousandth &&
            !accent1RedOff.HasGreenOffsetThousandth &&
            !accent1RedOff.HasBlueModulationThousandth &&
            !accent1RedOff.HasBlueOffsetThousandth &&
            !accent1RedOff.HasHueModulationThousandth &&
            !accent1RedOff.HasHueOffsetAngleThousandth &&
            !accent1RedOff.HasGray && !accent1RedOff.HasComp && !accent1RedOff.HasInv &&
            !accent1RedOff.HasGamma && !accent1RedOff.HasInvGamma &&
            accent1RedOff.RedOffsetThousandth >= -100_000 &&
            accent1RedOff.RedOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent1RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1RedOff", ["accentTransforms.accent1.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } greenModTheme &&
            greenModTheme.AccentTransforms[0] is { Role: "accent1", HasGreenModulationThousandth: true } accent1GreenMod &&
            !accent1GreenMod.HasTintThousandth &&
            !accent1GreenMod.HasShadeThousandth &&
            !accent1GreenMod.HasLuminanceModulationThousandth &&
            !accent1GreenMod.HasLuminanceOffsetThousandth &&
            !accent1GreenMod.HasAlphaModulationThousandth &&
            !accent1GreenMod.HasAlphaOffsetThousandth &&
            !accent1GreenMod.HasSaturationModulationThousandth &&
            !accent1GreenMod.HasSaturationOffsetThousandth &&
            !accent1GreenMod.HasRedModulationThousandth &&
            !accent1GreenMod.HasRedOffsetThousandth &&
            !accent1GreenMod.HasGreenOffsetThousandth &&
            !accent1GreenMod.HasBlueModulationThousandth &&
            !accent1GreenMod.HasBlueOffsetThousandth &&
            !accent1GreenMod.HasHueModulationThousandth &&
            !accent1GreenMod.HasHueOffsetAngleThousandth &&
            !accent1GreenMod.HasGray && !accent1GreenMod.HasComp && !accent1GreenMod.HasInv &&
            !accent1GreenMod.HasGamma && !accent1GreenMod.HasInvGamma &&
            accent1GreenMod.GreenModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent1GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1GreenMod", ["accentTransforms.accent1.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } greenOffTheme &&
            greenOffTheme.AccentTransforms[0] is { Role: "accent1", HasGreenOffsetThousandth: true } accent1GreenOff &&
            !accent1GreenOff.HasTintThousandth &&
            !accent1GreenOff.HasShadeThousandth &&
            !accent1GreenOff.HasLuminanceModulationThousandth &&
            !accent1GreenOff.HasLuminanceOffsetThousandth &&
            !accent1GreenOff.HasAlphaModulationThousandth &&
            !accent1GreenOff.HasAlphaOffsetThousandth &&
            !accent1GreenOff.HasSaturationModulationThousandth &&
            !accent1GreenOff.HasSaturationOffsetThousandth &&
            !accent1GreenOff.HasRedModulationThousandth &&
            !accent1GreenOff.HasRedOffsetThousandth &&
            !accent1GreenOff.HasGreenModulationThousandth &&
            !accent1GreenOff.HasBlueModulationThousandth &&
            !accent1GreenOff.HasBlueOffsetThousandth &&
            !accent1GreenOff.HasHueModulationThousandth &&
            !accent1GreenOff.HasHueOffsetAngleThousandth &&
            !accent1GreenOff.HasGray && !accent1GreenOff.HasComp && !accent1GreenOff.HasInv &&
            !accent1GreenOff.HasGamma && !accent1GreenOff.HasInvGamma &&
            accent1GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent1GreenOff.GreenOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent1GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1GreenOff", ["accentTransforms.accent1.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } blueModTheme &&
            blueModTheme.AccentTransforms[0] is { Role: "accent1", HasBlueModulationThousandth: true } accent1BlueMod &&
            !accent1BlueMod.HasTintThousandth &&
            !accent1BlueMod.HasShadeThousandth &&
            !accent1BlueMod.HasLuminanceModulationThousandth &&
            !accent1BlueMod.HasLuminanceOffsetThousandth &&
            !accent1BlueMod.HasAlphaModulationThousandth &&
            !accent1BlueMod.HasAlphaOffsetThousandth &&
            !accent1BlueMod.HasSaturationModulationThousandth &&
            !accent1BlueMod.HasSaturationOffsetThousandth &&
            !accent1BlueMod.HasRedModulationThousandth &&
            !accent1BlueMod.HasRedOffsetThousandth &&
            !accent1BlueMod.HasGreenModulationThousandth &&
            !accent1BlueMod.HasGreenOffsetThousandth &&
            !accent1BlueMod.HasBlueOffsetThousandth &&
            !accent1BlueMod.HasHueModulationThousandth &&
            !accent1BlueMod.HasHueOffsetAngleThousandth &&
            !accent1BlueMod.HasGray && !accent1BlueMod.HasComp && !accent1BlueMod.HasInv &&
            !accent1BlueMod.HasGamma && !accent1BlueMod.HasInvGamma &&
            accent1BlueMod.BlueModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent1BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1BlueMod", ["accentTransforms.accent1.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } blueOffTheme &&
            blueOffTheme.AccentTransforms[0] is { Role: "accent1", HasBlueOffsetThousandth: true } accent1BlueOff &&
            !accent1BlueOff.HasTintThousandth &&
            !accent1BlueOff.HasShadeThousandth &&
            !accent1BlueOff.HasLuminanceModulationThousandth &&
            !accent1BlueOff.HasLuminanceOffsetThousandth &&
            !accent1BlueOff.HasAlphaModulationThousandth &&
            !accent1BlueOff.HasAlphaOffsetThousandth &&
            !accent1BlueOff.HasSaturationModulationThousandth &&
            !accent1BlueOff.HasSaturationOffsetThousandth &&
            !accent1BlueOff.HasRedModulationThousandth &&
            !accent1BlueOff.HasRedOffsetThousandth &&
            !accent1BlueOff.HasGreenModulationThousandth &&
            !accent1BlueOff.HasGreenOffsetThousandth &&
            !accent1BlueOff.HasBlueModulationThousandth &&
            !accent1BlueOff.HasHueModulationThousandth &&
            !accent1BlueOff.HasHueOffsetAngleThousandth &&
            !accent1BlueOff.HasGray && !accent1BlueOff.HasComp && !accent1BlueOff.HasInv &&
            !accent1BlueOff.HasGamma && !accent1BlueOff.HasInvGamma &&
            accent1BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent1BlueOff.BlueOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent1BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1BlueOff", ["accentTransforms.accent1.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } hueModTheme &&
            hueModTheme.AccentTransforms[0] is { Role: "accent1", HasHueModulationThousandth: true } accent1HueMod &&
            !accent1HueMod.HasTintThousandth &&
            !accent1HueMod.HasShadeThousandth &&
            !accent1HueMod.HasLuminanceModulationThousandth &&
            !accent1HueMod.HasLuminanceOffsetThousandth &&
            !accent1HueMod.HasAlphaModulationThousandth &&
            !accent1HueMod.HasAlphaOffsetThousandth &&
            !accent1HueMod.HasSaturationModulationThousandth &&
            !accent1HueMod.HasSaturationOffsetThousandth &&
            !accent1HueMod.HasRedModulationThousandth &&
            !accent1HueMod.HasRedOffsetThousandth &&
            !accent1HueMod.HasGreenModulationThousandth &&
            !accent1HueMod.HasGreenOffsetThousandth &&
            !accent1HueMod.HasBlueModulationThousandth &&
            !accent1HueMod.HasBlueOffsetThousandth &&
            !accent1HueMod.HasHueOffsetAngleThousandth &&
            !accent1HueMod.HasGray && !accent1HueMod.HasComp && !accent1HueMod.HasInv &&
            !accent1HueMod.HasGamma && !accent1HueMod.HasInvGamma &&
            accent1HueMod.HueModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent1HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1HueMod", ["accentTransforms.accent1.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } hueOffTheme &&
            hueOffTheme.AccentTransforms[0] is { Role: "accent1", HasHueOffsetAngleThousandth: true } accent1HueOff &&
            !accent1HueOff.HasTintThousandth &&
            !accent1HueOff.HasShadeThousandth &&
            !accent1HueOff.HasLuminanceModulationThousandth &&
            !accent1HueOff.HasLuminanceOffsetThousandth &&
            !accent1HueOff.HasAlphaModulationThousandth &&
            !accent1HueOff.HasAlphaOffsetThousandth &&
            !accent1HueOff.HasSaturationModulationThousandth &&
            !accent1HueOff.HasSaturationOffsetThousandth &&
            !accent1HueOff.HasRedModulationThousandth &&
            !accent1HueOff.HasRedOffsetThousandth &&
            !accent1HueOff.HasGreenModulationThousandth &&
            !accent1HueOff.HasGreenOffsetThousandth &&
            !accent1HueOff.HasBlueModulationThousandth &&
            !accent1HueOff.HasBlueOffsetThousandth &&
            !accent1HueOff.HasHueModulationThousandth &&
            !accent1HueOff.HasGray && !accent1HueOff.HasComp && !accent1HueOff.HasInv &&
            !accent1HueOff.HasGamma && !accent1HueOff.HasInvGamma &&
            accent1HueOff.HueOffsetAngleThousandth >= -21_600_000 &&
            accent1HueOff.HueOffsetAngleThousandth <= 21_600_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent1"] = new JsonObject
                {
                    ["hueOff"] = JsonValue.Create(accent1HueOff.HueOffsetAngleThousandth / 60_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent1HueOff", ["accentTransforms.accent1.hueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2TintTheme &&
            accent2TintTheme.AccentTransforms[0] is { Role: "accent2", HasTintThousandth: true } accent2Tint &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent2Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2Tint", ["accentTransforms.accent2.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2ShadeTheme &&
            accent2ShadeTheme.AccentTransforms[0] is { Role: "accent2", HasShadeThousandth: true } accent2Shade &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent2Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2Shade", ["accentTransforms.accent2.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2LumModTheme &&
            accent2LumModTheme.AccentTransforms[0] is { Role: "accent2", HasLuminanceModulationThousandth: true } accent2LumMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent2LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2LumMod", ["accentTransforms.accent2.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2LumOffTheme &&
            accent2LumOffTheme.AccentTransforms[0] is { Role: "accent2", HasLuminanceOffsetThousandth: true } accent2LumOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent2LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2LumOff", ["accentTransforms.accent2.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2AlphaModTheme &&
            accent2AlphaModTheme.AccentTransforms[0] is { Role: "accent2", HasAlphaModulationThousandth: true } accent2AlphaMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent2AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2AlphaMod", ["accentTransforms.accent2.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2AlphaOffTheme &&
            accent2AlphaOffTheme.AccentTransforms[0] is { Role: "accent2", HasAlphaOffsetThousandth: true } accent2AlphaOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent2AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2AlphaOff", ["accentTransforms.accent2.alphaOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2SatModTheme &&
            accent2SatModTheme.AccentTransforms[0] is { Role: "accent2", HasSaturationModulationThousandth: true } accent2SatMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent2SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2SatMod", ["accentTransforms.accent2.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2SatOffTheme &&
            accent2SatOffTheme.AccentTransforms[0] is { Role: "accent2", HasSaturationOffsetThousandth: true } accent2SatOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent2SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2SatOff", ["accentTransforms.accent2.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2RedModTheme &&
            accent2RedModTheme.AccentTransforms[0] is { Role: "accent2", HasRedModulationThousandth: true } accent2RedMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent2RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2RedMod", ["accentTransforms.accent2.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2RedOffTheme &&
            accent2RedOffTheme.AccentTransforms[0] is { Role: "accent2", HasRedOffsetThousandth: true } accent2RedOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent2RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2RedOff", ["accentTransforms.accent2.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2GreenModTheme &&
            accent2GreenModTheme.AccentTransforms[0] is { Role: "accent2", HasGreenModulationThousandth: true } accent2GreenMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent2GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2GreenMod", ["accentTransforms.accent2.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2GreenOffTheme &&
            accent2GreenOffTheme.AccentTransforms[0] is { Role: "accent2", HasGreenOffsetThousandth: true } accent2GreenOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent2GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2GreenOff", ["accentTransforms.accent2.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2BlueModTheme &&
            accent2BlueModTheme.AccentTransforms[0] is { Role: "accent2", HasBlueModulationThousandth: true } accent2BlueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent2BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2BlueMod", ["accentTransforms.accent2.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2BlueOffTheme &&
            accent2BlueOffTheme.AccentTransforms[0] is { Role: "accent2", HasBlueOffsetThousandth: true } accent2BlueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent2BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2BlueOff", ["accentTransforms.accent2.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2HueModTheme &&
            accent2HueModTheme.AccentTransforms[0] is { Role: "accent2", HasHueModulationThousandth: true } accent2HueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent2HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2HueMod", ["accentTransforms.accent2.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent2HueOffTheme &&
            accent2HueOffTheme.AccentTransforms[0] is { Role: "accent2", HasHueOffsetAngleThousandth: true } accent2HueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent2"] = new JsonObject
                {
                    ["hueOff"] = JsonValue.Create(accent2HueOff.HueOffsetAngleThousandth / 60_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent2HueOff", ["accentTransforms.accent2.hueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3TintTheme &&
            accent3TintTheme.AccentTransforms[0] is { Role: "accent3", HasTintThousandth: true } accent3Tint &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent3Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3Tint", ["accentTransforms.accent3.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3ShadeTheme &&
            accent3ShadeTheme.AccentTransforms[0] is { Role: "accent3", HasShadeThousandth: true } accent3Shade &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent3Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3Shade", ["accentTransforms.accent3.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3LumModTheme &&
            accent3LumModTheme.AccentTransforms[0] is { Role: "accent3", HasLuminanceModulationThousandth: true } accent3LumMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent3LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3LumMod", ["accentTransforms.accent3.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5LumModTheme &&
            accent5LumModTheme.AccentTransforms[0] is { Role: "accent5", HasLuminanceModulationThousandth: true } accent5LumMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent5LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5LumMod", ["accentTransforms.accent5.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4LumModTheme &&
            accent4LumModTheme.AccentTransforms[0] is { Role: "accent4", HasLuminanceModulationThousandth: true } accent4LumMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent4LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4LumMod", ["accentTransforms.accent4.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5LumOffTheme &&
            accent5LumOffTheme.AccentTransforms[0] is { Role: "accent5", HasLuminanceOffsetThousandth: true } accent5LumOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent5LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5LumOff", ["accentTransforms.accent5.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5AlphaModTheme &&
            accent5AlphaModTheme.AccentTransforms[0] is { Role: "accent5", HasAlphaModulationThousandth: true } accent5AlphaMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent5AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5AlphaMod", ["accentTransforms.accent5.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5AlphaOffTheme &&
            accent5AlphaOffTheme.AccentTransforms[0] is { Role: "accent5", HasAlphaOffsetThousandth: true } accent5AlphaOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent5AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5AlphaOff", ["accentTransforms.accent5.alphaOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5SatModTheme &&
            accent5SatModTheme.AccentTransforms[0] is { Role: "accent5", HasSaturationModulationThousandth: true } accent5SatMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent5SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5SatMod", ["accentTransforms.accent5.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5SatOffTheme &&
            accent5SatOffTheme.AccentTransforms[0] is { Role: "accent5", HasSaturationOffsetThousandth: true } accent5SatOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent5SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5SatOff", ["accentTransforms.accent5.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5RedModTheme &&
            accent5RedModTheme.AccentTransforms[0] is { Role: "accent5", HasRedModulationThousandth: true } accent5RedMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent5RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5RedMod", ["accentTransforms.accent5.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5RedOffTheme &&
            accent5RedOffTheme.AccentTransforms[0] is { Role: "accent5", HasRedOffsetThousandth: true } accent5RedOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent5RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5RedOff", ["accentTransforms.accent5.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5GreenModTheme &&
            accent5GreenModTheme.AccentTransforms[0] is { Role: "accent5", HasGreenModulationThousandth: true } accent5GreenMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent5GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5GreenMod", ["accentTransforms.accent5.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5GreenOffTheme &&
            accent5GreenOffTheme.AccentTransforms[0] is { Role: "accent5", HasGreenOffsetThousandth: true } accent5GreenOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent5GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5GreenOff", ["accentTransforms.accent5.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5BlueModTheme &&
            accent5BlueModTheme.AccentTransforms[0] is { Role: "accent5", HasBlueModulationThousandth: true } accent5BlueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent5BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5BlueMod", ["accentTransforms.accent5.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5BlueOffTheme &&
            accent5BlueOffTheme.AccentTransforms[0] is { Role: "accent5", HasBlueOffsetThousandth: true } accent5BlueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent5BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5BlueOff", ["accentTransforms.accent5.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5HueModTheme &&
            accent5HueModTheme.AccentTransforms[0] is { Role: "accent5", HasHueModulationThousandth: true } accent5HueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent5HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5HueMod", ["accentTransforms.accent5.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5HueOffTheme &&
            accent5HueOffTheme.AccentTransforms[0] is { Role: "accent5", HasHueOffsetAngleThousandth: true } accent5HueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["hueOff"] = JsonValue.Create(accent5HueOff.HueOffsetAngleThousandth / 60_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5HueOff", ["accentTransforms.accent5.hueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4LumOffTheme &&
            accent4LumOffTheme.AccentTransforms[0] is { Role: "accent4", HasLuminanceOffsetThousandth: true } accent4LumOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent4LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4LumOff", ["accentTransforms.accent4.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4AlphaModTheme &&
            accent4AlphaModTheme.AccentTransforms[0] is { Role: "accent4", HasAlphaModulationThousandth: true } accent4AlphaMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent4AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4AlphaMod", ["accentTransforms.accent4.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4AlphaOffTheme &&
            accent4AlphaOffTheme.AccentTransforms[0] is { Role: "accent4", HasAlphaOffsetThousandth: true } accent4AlphaOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent4AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4AlphaOff", ["accentTransforms.accent4.alphaOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4SatModTheme &&
            accent4SatModTheme.AccentTransforms[0] is { Role: "accent4", HasSaturationModulationThousandth: true } accent4SatMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent4SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4SatMod", ["accentTransforms.accent4.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4SatOffTheme &&
            accent4SatOffTheme.AccentTransforms[0] is { Role: "accent4", HasSaturationOffsetThousandth: true } accent4SatOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent4SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4SatOff", ["accentTransforms.accent4.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4RedModTheme &&
            accent4RedModTheme.AccentTransforms[0] is { Role: "accent4", HasRedModulationThousandth: true } accent4RedMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent4RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4RedMod", ["accentTransforms.accent4.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4GreenModTheme &&
            accent4GreenModTheme.AccentTransforms[0] is { Role: "accent4", HasGreenModulationThousandth: true } accent4GreenMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent4GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4GreenMod", ["accentTransforms.accent4.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4BlueModTheme &&
            accent4BlueModTheme.AccentTransforms[0] is { Role: "accent4", HasBlueModulationThousandth: true } accent4BlueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent4BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4BlueMod", ["accentTransforms.accent4.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4HueModTheme &&
            accent4HueModTheme.AccentTransforms[0] is { Role: "accent4", HasHueModulationThousandth: true } accent4HueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent4HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4HueMod", ["accentTransforms.accent4.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4BlueOffTheme &&
            accent4BlueOffTheme.AccentTransforms[0] is { Role: "accent4", HasBlueOffsetThousandth: true } accent4BlueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent4BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4BlueOff", ["accentTransforms.accent4.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4GreenOffTheme &&
            accent4GreenOffTheme.AccentTransforms[0] is { Role: "accent4", HasGreenOffsetThousandth: true } accent4GreenOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent4GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4GreenOff", ["accentTransforms.accent4.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4RedOffTheme &&
            accent4RedOffTheme.AccentTransforms[0] is { Role: "accent4", HasRedOffsetThousandth: true } accent4RedOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent4RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4RedOff", ["accentTransforms.accent4.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3LumOffTheme &&
            accent3LumOffTheme.AccentTransforms[0] is { Role: "accent3", HasLuminanceOffsetThousandth: true } accent3LumOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent3LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3LumOff", ["accentTransforms.accent3.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3AlphaModTheme &&
            accent3AlphaModTheme.AccentTransforms[0] is { Role: "accent3", HasAlphaModulationThousandth: true } accent3AlphaMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent3AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3AlphaMod", ["accentTransforms.accent3.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3SatModTheme &&
            accent3SatModTheme.AccentTransforms[0] is { Role: "accent3", HasSaturationModulationThousandth: true } accent3SatMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent3SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3SatMod", ["accentTransforms.accent3.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3SatOffTheme &&
            accent3SatOffTheme.AccentTransforms[0] is { Role: "accent3", HasSaturationOffsetThousandth: true } accent3SatOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent3SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3SatOff", ["accentTransforms.accent3.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3RedModTheme &&
            accent3RedModTheme.AccentTransforms[0] is { Role: "accent3", HasRedModulationThousandth: true } accent3RedMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent3RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3RedMod", ["accentTransforms.accent3.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3RedOffTheme &&
            accent3RedOffTheme.AccentTransforms[0] is { Role: "accent3", HasRedOffsetThousandth: true } accent3RedOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent3RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3RedOff", ["accentTransforms.accent3.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3GreenModTheme &&
            accent3GreenModTheme.AccentTransforms[0] is { Role: "accent3", HasGreenModulationThousandth: true } accent3GreenMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent3GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3GreenMod", ["accentTransforms.accent3.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3GreenOffTheme &&
            accent3GreenOffTheme.AccentTransforms[0] is { Role: "accent3", HasGreenOffsetThousandth: true } accent3GreenOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent3GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3GreenOff", ["accentTransforms.accent3.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3BlueModTheme &&
            accent3BlueModTheme.AccentTransforms[0] is { Role: "accent3", HasBlueModulationThousandth: true } accent3BlueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent3BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3BlueMod", ["accentTransforms.accent3.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3BlueOffTheme &&
            accent3BlueOffTheme.AccentTransforms[0] is { Role: "accent3", HasBlueOffsetThousandth: true } accent3BlueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent3BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3BlueOff", ["accentTransforms.accent3.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3HueModTheme &&
            accent3HueModTheme.AccentTransforms[0] is { Role: "accent3", HasHueModulationThousandth: true } accent3HueMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent3HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3HueMod", ["accentTransforms.accent3.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3HueOffTheme &&
            accent3HueOffTheme.AccentTransforms[0] is { Role: "accent3", HasHueOffsetAngleThousandth: true } accent3HueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["hueOff"] = JsonValue.Create(accent3HueOff.HueOffsetAngleThousandth / 60_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3HueOff", ["accentTransforms.accent3.hueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4HueOffTheme &&
            accent4HueOffTheme.AccentTransforms[0] is { Role: "accent4", HasHueOffsetAngleThousandth: true } accent4HueOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["hueOff"] = JsonValue.Create(accent4HueOff.HueOffsetAngleThousandth / 60_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4HueOff", ["accentTransforms.accent4.hueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5TintTheme &&
            accent5TintTheme.AccentTransforms[0] is { Role: "accent5", HasTintThousandth: true } accent5Tint &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent5Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5Tint", ["accentTransforms.accent5.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6TintTheme &&
            accent6TintTheme.AccentTransforms[0] is { Role: "accent6", HasTintThousandth: true } accent6Tint &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent6Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6Tint", ["accentTransforms.accent6.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6ShadeTheme &&
            accent6ShadeTheme.AccentTransforms[0] is { Role: "accent6", HasShadeThousandth: true } accent6Shade &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent6Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6Shade", ["accentTransforms.accent6.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6LumModTheme &&
            accent6LumModTheme.AccentTransforms[0] is { Role: "accent6", HasLuminanceModulationThousandth: true } accent6LumMod &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["lumMod"] = JsonValue.Create(accent6LumMod.LuminanceModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6LumMod", ["accentTransforms.accent6.lumMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6LumOffTheme &&
            accent6LumOffTheme.AccentTransforms[0] is { Role: "accent6", HasLuminanceOffsetThousandth: true } accent6LumOff &&
            !accent6LumOff.HasTintThousandth &&
            !accent6LumOff.HasShadeThousandth &&
            !accent6LumOff.HasLuminanceModulationThousandth &&
            !accent6LumOff.HasAlphaModulationThousandth &&
            !accent6LumOff.HasAlphaOffsetThousandth &&
            !accent6LumOff.HasSaturationModulationThousandth &&
            !accent6LumOff.HasSaturationOffsetThousandth &&
            !accent6LumOff.HasRedModulationThousandth &&
            !accent6LumOff.HasRedOffsetThousandth &&
            !accent6LumOff.HasGreenModulationThousandth &&
            !accent6LumOff.HasGreenOffsetThousandth &&
            !accent6LumOff.HasBlueModulationThousandth &&
            !accent6LumOff.HasBlueOffsetThousandth &&
            !accent6LumOff.HasHueModulationThousandth &&
            !accent6LumOff.HasHueOffsetAngleThousandth &&
            !accent6LumOff.HasGray && !accent6LumOff.HasComp && !accent6LumOff.HasInv &&
            !accent6LumOff.HasGamma && !accent6LumOff.HasInvGamma &&
            accent6LumOff.LuminanceOffsetThousandth >= -100_000 &&
            accent6LumOff.LuminanceOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["lumOff"] = JsonValue.Create(accent6LumOff.LuminanceOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6LumOff", ["accentTransforms.accent6.lumOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6AlphaModTheme &&
            accent6AlphaModTheme.AccentTransforms[0] is { Role: "accent6", HasAlphaModulationThousandth: true } accent6AlphaMod &&
            !accent6AlphaMod.HasTintThousandth &&
            !accent6AlphaMod.HasShadeThousandth &&
            !accent6AlphaMod.HasLuminanceModulationThousandth &&
            !accent6AlphaMod.HasLuminanceOffsetThousandth &&
            !accent6AlphaMod.HasAlphaOffsetThousandth &&
            !accent6AlphaMod.HasSaturationModulationThousandth &&
            !accent6AlphaMod.HasSaturationOffsetThousandth &&
            !accent6AlphaMod.HasRedModulationThousandth &&
            !accent6AlphaMod.HasRedOffsetThousandth &&
            !accent6AlphaMod.HasGreenModulationThousandth &&
            !accent6AlphaMod.HasGreenOffsetThousandth &&
            !accent6AlphaMod.HasBlueModulationThousandth &&
            !accent6AlphaMod.HasBlueOffsetThousandth &&
            !accent6AlphaMod.HasHueModulationThousandth &&
            !accent6AlphaMod.HasHueOffsetAngleThousandth &&
            !accent6AlphaMod.HasGray && !accent6AlphaMod.HasComp && !accent6AlphaMod.HasInv &&
            !accent6AlphaMod.HasGamma && !accent6AlphaMod.HasInvGamma &&
            accent6AlphaMod.AlphaModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["alphaMod"] = JsonValue.Create(accent6AlphaMod.AlphaModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6AlphaMod", ["accentTransforms.accent6.alphaMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6AlphaOffTheme &&
            accent6AlphaOffTheme.AccentTransforms[0] is { Role: "accent6", HasAlphaOffsetThousandth: true } accent6AlphaOff &&
            !accent6AlphaOff.HasTintThousandth &&
            !accent6AlphaOff.HasShadeThousandth &&
            !accent6AlphaOff.HasLuminanceModulationThousandth &&
            !accent6AlphaOff.HasLuminanceOffsetThousandth &&
            !accent6AlphaOff.HasAlphaModulationThousandth &&
            !accent6AlphaOff.HasSaturationModulationThousandth &&
            !accent6AlphaOff.HasSaturationOffsetThousandth &&
            !accent6AlphaOff.HasRedModulationThousandth &&
            !accent6AlphaOff.HasRedOffsetThousandth &&
            !accent6AlphaOff.HasGreenModulationThousandth &&
            !accent6AlphaOff.HasGreenOffsetThousandth &&
            !accent6AlphaOff.HasBlueModulationThousandth &&
            !accent6AlphaOff.HasBlueOffsetThousandth &&
            !accent6AlphaOff.HasHueModulationThousandth &&
            !accent6AlphaOff.HasHueOffsetAngleThousandth &&
            !accent6AlphaOff.HasGray && !accent6AlphaOff.HasComp && !accent6AlphaOff.HasInv &&
            !accent6AlphaOff.HasGamma && !accent6AlphaOff.HasInvGamma &&
            accent6AlphaOff.AlphaOffsetThousandth >= -100_000 &&
            accent6AlphaOff.AlphaOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent6AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6AlphaOff", ["accentTransforms.accent6.alphaOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6SatModTheme &&
            accent6SatModTheme.AccentTransforms[0] is { Role: "accent6", HasSaturationModulationThousandth: true } accent6SatMod &&
            !accent6SatMod.HasTintThousandth &&
            !accent6SatMod.HasShadeThousandth &&
            !accent6SatMod.HasLuminanceModulationThousandth &&
            !accent6SatMod.HasLuminanceOffsetThousandth &&
            !accent6SatMod.HasAlphaModulationThousandth &&
            !accent6SatMod.HasAlphaOffsetThousandth &&
            !accent6SatMod.HasSaturationOffsetThousandth &&
            !accent6SatMod.HasRedModulationThousandth &&
            !accent6SatMod.HasRedOffsetThousandth &&
            !accent6SatMod.HasGreenModulationThousandth &&
            !accent6SatMod.HasGreenOffsetThousandth &&
            !accent6SatMod.HasBlueModulationThousandth &&
            !accent6SatMod.HasBlueOffsetThousandth &&
            !accent6SatMod.HasHueModulationThousandth &&
            !accent6SatMod.HasHueOffsetAngleThousandth &&
            !accent6SatMod.HasGray && !accent6SatMod.HasComp && !accent6SatMod.HasInv &&
            !accent6SatMod.HasGamma && !accent6SatMod.HasInvGamma &&
            accent6SatMod.SaturationModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["satMod"] = JsonValue.Create(accent6SatMod.SaturationModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6SatMod", ["accentTransforms.accent6.satMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6SatOffTheme &&
            accent6SatOffTheme.AccentTransforms[0] is { Role: "accent6", HasSaturationOffsetThousandth: true } accent6SatOff &&
            !accent6SatOff.HasTintThousandth &&
            !accent6SatOff.HasShadeThousandth &&
            !accent6SatOff.HasLuminanceModulationThousandth &&
            !accent6SatOff.HasLuminanceOffsetThousandth &&
            !accent6SatOff.HasAlphaModulationThousandth &&
            !accent6SatOff.HasAlphaOffsetThousandth &&
            !accent6SatOff.HasSaturationModulationThousandth &&
            !accent6SatOff.HasRedModulationThousandth &&
            !accent6SatOff.HasRedOffsetThousandth &&
            !accent6SatOff.HasGreenModulationThousandth &&
            !accent6SatOff.HasGreenOffsetThousandth &&
            !accent6SatOff.HasBlueModulationThousandth &&
            !accent6SatOff.HasBlueOffsetThousandth &&
            !accent6SatOff.HasHueModulationThousandth &&
            !accent6SatOff.HasHueOffsetAngleThousandth &&
            !accent6SatOff.HasGray && !accent6SatOff.HasComp && !accent6SatOff.HasInv &&
            !accent6SatOff.HasGamma && !accent6SatOff.HasInvGamma &&
            accent6SatOff.SaturationOffsetThousandth >= -100_000 &&
            accent6SatOff.SaturationOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["satOff"] = JsonValue.Create(accent6SatOff.SaturationOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6SatOff", ["accentTransforms.accent6.satOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6RedModTheme &&
            accent6RedModTheme.AccentTransforms[0] is { Role: "accent6", HasRedModulationThousandth: true } accent6RedMod &&
            !accent6RedMod.HasTintThousandth &&
            !accent6RedMod.HasShadeThousandth &&
            !accent6RedMod.HasLuminanceModulationThousandth &&
            !accent6RedMod.HasLuminanceOffsetThousandth &&
            !accent6RedMod.HasAlphaModulationThousandth &&
            !accent6RedMod.HasAlphaOffsetThousandth &&
            !accent6RedMod.HasSaturationModulationThousandth &&
            !accent6RedMod.HasSaturationOffsetThousandth &&
            !accent6RedMod.HasRedOffsetThousandth &&
            !accent6RedMod.HasGreenModulationThousandth &&
            !accent6RedMod.HasGreenOffsetThousandth &&
            !accent6RedMod.HasBlueModulationThousandth &&
            !accent6RedMod.HasBlueOffsetThousandth &&
            !accent6RedMod.HasHueModulationThousandth &&
            !accent6RedMod.HasHueOffsetAngleThousandth &&
            !accent6RedMod.HasGray && !accent6RedMod.HasComp && !accent6RedMod.HasInv &&
            !accent6RedMod.HasGamma && !accent6RedMod.HasInvGamma &&
            accent6RedMod.RedModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["redMod"] = JsonValue.Create(accent6RedMod.RedModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6RedMod", ["accentTransforms.accent6.redMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6RedOffTheme &&
            accent6RedOffTheme.AccentTransforms[0] is { Role: "accent6", HasRedOffsetThousandth: true } accent6RedOff &&
            !accent6RedOff.HasTintThousandth &&
            !accent6RedOff.HasShadeThousandth &&
            !accent6RedOff.HasLuminanceModulationThousandth &&
            !accent6RedOff.HasLuminanceOffsetThousandth &&
            !accent6RedOff.HasAlphaModulationThousandth &&
            !accent6RedOff.HasAlphaOffsetThousandth &&
            !accent6RedOff.HasSaturationModulationThousandth &&
            !accent6RedOff.HasSaturationOffsetThousandth &&
            !accent6RedOff.HasRedModulationThousandth &&
            !accent6RedOff.HasGreenModulationThousandth &&
            !accent6RedOff.HasGreenOffsetThousandth &&
            !accent6RedOff.HasBlueModulationThousandth &&
            !accent6RedOff.HasBlueOffsetThousandth &&
            !accent6RedOff.HasHueModulationThousandth &&
            !accent6RedOff.HasHueOffsetAngleThousandth &&
            !accent6RedOff.HasGray && !accent6RedOff.HasComp && !accent6RedOff.HasInv &&
            !accent6RedOff.HasGamma && !accent6RedOff.HasInvGamma &&
            accent6RedOff.RedOffsetThousandth >= -100_000 &&
            accent6RedOff.RedOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["redOff"] = JsonValue.Create(accent6RedOff.RedOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6RedOff", ["accentTransforms.accent6.redOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6GreenModTheme &&
            accent6GreenModTheme.AccentTransforms[0] is { Role: "accent6", HasGreenModulationThousandth: true } accent6GreenMod &&
            !accent6GreenMod.HasTintThousandth &&
            !accent6GreenMod.HasShadeThousandth &&
            !accent6GreenMod.HasLuminanceModulationThousandth &&
            !accent6GreenMod.HasLuminanceOffsetThousandth &&
            !accent6GreenMod.HasAlphaModulationThousandth &&
            !accent6GreenMod.HasAlphaOffsetThousandth &&
            !accent6GreenMod.HasSaturationModulationThousandth &&
            !accent6GreenMod.HasSaturationOffsetThousandth &&
            !accent6GreenMod.HasRedOffsetThousandth &&
            !accent6GreenMod.HasGreenOffsetThousandth &&
            !accent6GreenMod.HasBlueModulationThousandth &&
            !accent6GreenMod.HasBlueOffsetThousandth &&
            !accent6GreenMod.HasHueModulationThousandth &&
            !accent6GreenMod.HasHueOffsetAngleThousandth &&
            !accent6GreenMod.HasGray && !accent6GreenMod.HasComp && !accent6GreenMod.HasInv &&
            !accent6GreenMod.HasGamma && !accent6GreenMod.HasInvGamma &&
            accent6GreenMod.GreenModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["greenMod"] = JsonValue.Create(accent6GreenMod.GreenModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6GreenMod", ["accentTransforms.accent6.greenMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6GreenOffTheme &&
            accent6GreenOffTheme.AccentTransforms[0] is { Role: "accent6", HasGreenOffsetThousandth: true } accent6GreenOff &&
            !accent6GreenOff.HasTintThousandth &&
            !accent6GreenOff.HasShadeThousandth &&
            !accent6GreenOff.HasLuminanceModulationThousandth &&
            !accent6GreenOff.HasLuminanceOffsetThousandth &&
            !accent6GreenOff.HasAlphaModulationThousandth &&
            !accent6GreenOff.HasAlphaOffsetThousandth &&
            !accent6GreenOff.HasSaturationModulationThousandth &&
            !accent6GreenOff.HasSaturationOffsetThousandth &&
            !accent6GreenOff.HasRedOffsetThousandth &&
            !accent6GreenOff.HasGreenModulationThousandth &&
            !accent6GreenOff.HasBlueModulationThousandth &&
            !accent6GreenOff.HasBlueOffsetThousandth &&
            !accent6GreenOff.HasHueModulationThousandth &&
            !accent6GreenOff.HasHueOffsetAngleThousandth &&
            !accent6GreenOff.HasGray && !accent6GreenOff.HasComp && !accent6GreenOff.HasInv &&
            !accent6GreenOff.HasGamma && !accent6GreenOff.HasInvGamma &&
            accent6GreenOff.GreenOffsetThousandth >= -100_000 &&
            accent6GreenOff.GreenOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["greenOff"] = JsonValue.Create(accent6GreenOff.GreenOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6GreenOff", ["accentTransforms.accent6.greenOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6BlueModTheme &&
            accent6BlueModTheme.AccentTransforms[0] is { Role: "accent6", HasBlueModulationThousandth: true } accent6BlueMod &&
            !accent6BlueMod.HasTintThousandth &&
            !accent6BlueMod.HasShadeThousandth &&
            !accent6BlueMod.HasLuminanceModulationThousandth &&
            !accent6BlueMod.HasLuminanceOffsetThousandth &&
            !accent6BlueMod.HasAlphaModulationThousandth &&
            !accent6BlueMod.HasAlphaOffsetThousandth &&
            !accent6BlueMod.HasSaturationModulationThousandth &&
            !accent6BlueMod.HasSaturationOffsetThousandth &&
            !accent6BlueMod.HasRedOffsetThousandth &&
            !accent6BlueMod.HasGreenOffsetThousandth &&
            !accent6BlueMod.HasGreenModulationThousandth &&
            !accent6BlueMod.HasBlueOffsetThousandth &&
            !accent6BlueMod.HasHueModulationThousandth &&
            !accent6BlueMod.HasHueOffsetAngleThousandth &&
            !accent6BlueMod.HasGray && !accent6BlueMod.HasComp && !accent6BlueMod.HasInv &&
            !accent6BlueMod.HasGamma && !accent6BlueMod.HasInvGamma &&
            accent6BlueMod.BlueModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["blueMod"] = JsonValue.Create(accent6BlueMod.BlueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6BlueMod", ["accentTransforms.accent6.blueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6BlueOffTheme &&
            accent6BlueOffTheme.AccentTransforms[0] is { Role: "accent6", HasBlueOffsetThousandth: true } accent6BlueOff &&
            !accent6BlueOff.HasTintThousandth &&
            !accent6BlueOff.HasShadeThousandth &&
            !accent6BlueOff.HasLuminanceModulationThousandth &&
            !accent6BlueOff.HasLuminanceOffsetThousandth &&
            !accent6BlueOff.HasAlphaModulationThousandth &&
            !accent6BlueOff.HasAlphaOffsetThousandth &&
            !accent6BlueOff.HasSaturationModulationThousandth &&
            !accent6BlueOff.HasSaturationOffsetThousandth &&
            !accent6BlueOff.HasRedOffsetThousandth &&
            !accent6BlueOff.HasGreenModulationThousandth &&
            !accent6BlueOff.HasBlueModulationThousandth &&
            !accent6BlueOff.HasGreenOffsetThousandth &&
            !accent6BlueOff.HasHueModulationThousandth &&
            !accent6BlueOff.HasHueOffsetAngleThousandth &&
            !accent6BlueOff.HasGray && !accent6BlueOff.HasComp && !accent6BlueOff.HasInv &&
            !accent6BlueOff.HasGamma && !accent6BlueOff.HasInvGamma &&
            accent6BlueOff.BlueOffsetThousandth >= -100_000 &&
            accent6BlueOff.BlueOffsetThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["blueOff"] = JsonValue.Create(accent6BlueOff.BlueOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6BlueOff", ["accentTransforms.accent6.blueOff"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent6HueModTheme &&
            accent6HueModTheme.AccentTransforms[0] is { Role: "accent6", HasHueModulationThousandth: true } accent6HueMod &&
            !accent6HueMod.HasTintThousandth &&
            !accent6HueMod.HasShadeThousandth &&
            !accent6HueMod.HasLuminanceModulationThousandth &&
            !accent6HueMod.HasLuminanceOffsetThousandth &&
            !accent6HueMod.HasAlphaModulationThousandth &&
            !accent6HueMod.HasAlphaOffsetThousandth &&
            !accent6HueMod.HasSaturationModulationThousandth &&
            !accent6HueMod.HasSaturationOffsetThousandth &&
            !accent6HueMod.HasRedModulationThousandth &&
            !accent6HueMod.HasRedOffsetThousandth &&
            !accent6HueMod.HasGreenModulationThousandth &&
            !accent6HueMod.HasGreenOffsetThousandth &&
            !accent6HueMod.HasBlueModulationThousandth &&
            !accent6HueMod.HasHueOffsetAngleThousandth &&
            !accent6HueMod.HasGray && !accent6HueMod.HasComp && !accent6HueMod.HasInv &&
            !accent6HueMod.HasGamma && !accent6HueMod.HasInvGamma &&
            accent6HueMod.HueModulationThousandth <= 100_000)
        {
            theme["accentTransforms"] = new JsonObject
            {
                ["accent6"] = new JsonObject
                {
                    ["hueMod"] = JsonValue.Create(accent6HueMod.HueModulationThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent6HueMod", ["accentTransforms.accent6.hueMod"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4TintTheme &&
            accent4TintTheme.AccentTransforms[0] is { Role: "accent4", HasTintThousandth: true } accent4Tint &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["tint"] = JsonValue.Create(accent4Tint.TintThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4Tint", ["accentTransforms.accent4.tint"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent5ShadeTheme &&
            accent5ShadeTheme.AccentTransforms[0] is { Role: "accent5", HasShadeThousandth: true } accent5Shade &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent5"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent5Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent5Shade", ["accentTransforms.accent5.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent4ShadeTheme &&
            accent4ShadeTheme.AccentTransforms[0] is { Role: "accent4", HasShadeThousandth: true } accent4Shade &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent4"] = new JsonObject
                {
                    ["shade"] = JsonValue.Create(accent4Shade.ShadeThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent4Shade", ["accentTransforms.accent4.shade"]));
        }
        if (presentation.AuthoredTheme is { AccentTransforms.Count: 1 } accent3AlphaOffTheme &&
            accent3AlphaOffTheme.AccentTransforms[0] is { Role: "accent3", HasAlphaOffsetThousandth: true } accent3AlphaOff &&
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
            theme["accentTransforms"] = new JsonObject
            {
                ["accent3"] = new JsonObject
                {
                    ["alphaOff"] = JsonValue.Create(accent3AlphaOff.AlphaOffsetThousandth / 100_000d),
                },
            };
            themeCapabilities.Add(new("setThemeAccent3AlphaOff", ["accentTransforms.accent3.alphaOff"]));
        }
        if (presentation.AuthoredTheme?.HasMajorFontFamily == true &&
            presentation.AuthoredTheme.HasMinorFontFamily)
        {
            var fontScheme = new JsonObject
            {
                ["major"] = StringNode(presentation.AuthoredTheme.MajorFontFamily),
                ["minor"] = StringNode(presentation.AuthoredTheme.MinorFontFamily),
            };
            if (presentation.AuthoredTheme.HasMajorFontFamilyEastAsia)
                fontScheme["majorEastAsia"] = StringNode(presentation.AuthoredTheme.MajorFontFamilyEastAsia);
            if (presentation.AuthoredTheme.HasMajorFontFamilyComplexScript)
                fontScheme["majorComplexScript"] = StringNode(presentation.AuthoredTheme.MajorFontFamilyComplexScript);
            if (presentation.AuthoredTheme.HasMinorFontFamilyEastAsia)
                fontScheme["minorEastAsia"] = StringNode(presentation.AuthoredTheme.MinorFontFamilyEastAsia);
            if (presentation.AuthoredTheme.HasMinorFontFamilyComplexScript)
                fontScheme["minorComplexScript"] = StringNode(presentation.AuthoredTheme.MinorFontFamilyComplexScript);
            theme["fontScheme"] = fontScheme;
            themeCapabilities.Add(new("setThemeFontScheme", ["fontScheme.major"]));
            themeCapabilities.Add(new("setThemeMinorFont", ["fontScheme.minor"]));
            if (presentation.AuthoredTheme.HasMajorFontFamilyEastAsia)
                themeCapabilities.Add(new("setThemeMajorFontEastAsia", ["fontScheme.majorEastAsia"]));
            if (presentation.AuthoredTheme.HasMinorFontFamilyEastAsia)
                themeCapabilities.Add(new("setThemeMinorFontEastAsia", ["fontScheme.minorEastAsia"]));
            if (presentation.AuthoredTheme.HasMajorFontFamilyComplexScript)
                themeCapabilities.Add(new("setThemeMajorFontComplexScript", ["fontScheme.majorComplexScript"]));
            if (presentation.AuthoredTheme.HasMinorFontFamilyComplexScript)
                themeCapabilities.Add(new("setThemeMinorFontComplexScript", ["fontScheme.minorComplexScript"]));
        }
        if (themeCapabilities.Count > 0)
        {
            var themeObjectHash = Sha256(Encoding.UTF8.GetBytes("officekit:ppj:presentation-theme-name"));
            theme["nativeRef"] = NativeRef(
                context,
                "theme",
                themeObjectHash,
                themeCapabilities);
        }
        var output = new JsonObject
        {
            ["canvas"] = ProjectCanvas(presentation, context),
            ["theme"] = theme,
            ["fonts"] = new JsonArray(new JsonObject
            {
                ["id"] = StringNode("source-font"),
                ["family"] = StringNode("Arial"),
                ["language"] = StringNode("und"),
            }),
            ["styles"] = new JsonObject(),
            ["grammar"] = new JsonObject
            {
                ["name"] = StringNode("Source-derived projection"),
                ["rationale"] = StringNode("The original PPTX remains the authority for native visual styling that PPJ does not model."),
                ["visualThesis"] = StringNode("Preserve and continue the imported presentation without inventing its original intent."),
                ["surfaceHierarchy"] = new JsonArray { StringNode("Keep source-owned surfaces unchanged unless a capability explicitly permits an edit.") },
                ["typographyRhythm"] = new JsonArray { StringNode("Retain imported typography through the source package.") },
                ["geometryRules"] = new JsonArray { StringNode("Retain imported geometry and z-order unless a nativeRef capability permits a change.") },
                ["densityRhythm"] = new JsonArray { StringNode("Retain the source page density.") },
                ["carrierRules"] = new JsonArray { StringNode("Use projected typed objects where safe and opaque native objects everywhere else.") },
                ["forbiddenPatterns"] = new JsonArray { StringNode("Rebuilding opaque content"), StringNode("Guessing unsupported native semantics") },
            },
            ["motionPolicy"] = StringNode("explicit"),
        };
        if (presentation.Masters.Count > 0)
            output["masters"] = ProjectMasters(presentation.Masters, context);
        if (presentation.Layouts.Count > 0)
            output["layouts"] = ProjectLayouts(presentation.Layouts, presentation.Masters, context);
        return output;
    }

    private static JsonArray ProjectMasters(
        IEnumerable<PresentationMaster> masters,
        ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var master in masters)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.MasterId(master.Id)),
                ["name"] = StringNode(master.Name),
            };
            if (master.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                if (master.Source.BackgroundEditable)
                    capabilities.Add(new("setBackground", ["background"]));
                if (master.Source.TextStylesEditable)
                    capabilities.Add(new("setTextParagraphStyle", ["textStyles"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"master:{master.Id}",
                    HashOrFallback(master.Source.MasterXmlSha256, master),
                    capabilities);
            }
            if (ProjectBackground(master.Background, context) is { } background)
                projected["background"] = background;
            var textStyles = new JsonObject();
            AddMasterTextLevels(textStyles, "title", master.TextStyles?.TitleLevels, context);
            AddMasterTextLevels(textStyles, "body", master.TextStyles?.BodyLevels, context);
            AddMasterTextLevels(textStyles, "other", master.TextStyles?.OtherLevels, context);
            if (textStyles.Count > 0) projected["textStyles"] = textStyles;
            var placeholders = ProjectLayoutPlaceholders(master.Placeholders, null, context);
            if (placeholders.Count > 0) projected["placeholders"] = placeholders;
            output.Add(projected);
        }
        return output;
    }

    private static JsonArray ProjectLayouts(
        IEnumerable<PresentationLayout> layouts,
        IEnumerable<PresentationMaster> masters,
        ProjectionContext context)
    {
        var mastersById = masters.ToDictionary(master => master.Id, StringComparer.Ordinal);
        var output = new JsonArray();
        foreach (var layout in layouts)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.LayoutId(layout.Id)),
                ["name"] = StringNode(layout.Name),
                ["master"] = StringNode(context.MasterId(layout.MasterId)),
                ["layoutType"] = StringNode(layout.Type),
            };
            if (layout.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                if (layout.Source.BackgroundEditable)
                    capabilities.Add(new("setBackground", ["background"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"layout:{layout.Id}",
                    HashOrFallback(layout.Source.LayoutXmlSha256, layout),
                    capabilities);
            }
            if (ProjectBackground(layout.Background, context) is { } background)
                projected["background"] = background;
            var inheritedPlaceholders = mastersById.TryGetValue(layout.MasterId, out var master)
                ? master.Placeholders
                : null;
            var placeholders = ProjectLayoutPlaceholders(layout.Placeholders, inheritedPlaceholders, context);
            if (placeholders.Count > 0) projected["placeholders"] = placeholders;
            output.Add(projected);
        }
        return output;
    }

    private static JsonArray ProjectLayoutPlaceholders(
        IEnumerable<PresentationPlaceholder> placeholders,
        IEnumerable<PresentationPlaceholder>? inheritedPlaceholders,
        ProjectionContext context)
    {
        var inheritedFrames = inheritedPlaceholders is null
            ? new Dictionary<(string Type, uint Index), PresentationPlaceholderFrame>()
            : inheritedPlaceholders
                .Where(placeholder => placeholder.DirectFrame is not null)
                .GroupBy(placeholder => (placeholder.Type, placeholder.Index))
                .Where(group => group.Count() == 1)
                .ToDictionary(
                    group => group.Key,
                    group => group.Single().DirectFrame!.Clone());
        var output = new JsonArray();
        foreach (var placeholder in placeholders)
        {
            var frameSource = placeholder.DirectFrame;
            if (frameSource is null &&
                inheritedFrames.TryGetValue((placeholder.Type, placeholder.Index), out var inheritedFrame))
                frameSource = inheritedFrame;
            // An inherited placeholder without a matching master frame, or an
            // irregular transform that was intentionally rejected by the
            // native reader, still has no safe PPJ frame. Keep it source-owned
            // instead of inventing coordinates.
            if (frameSource is null) continue;
            var frame = new JsonObject
            {
                ["x"] = JsonValue.Create(Points(frameSource.LeftEmu)),
                ["y"] = JsonValue.Create(Points(frameSource.TopEmu)),
                ["width"] = JsonValue.Create(Math.Max(0.001, Points(frameSource.WidthEmu))),
                ["height"] = JsonValue.Create(Math.Max(0.001, Points(frameSource.HeightEmu))),
            };
            if (frameSource.HasRotationAngle60000)
                frame["rotation"] = JsonValue.Create(frameSource.RotationAngle60000 / 60_000d);
            if (frameSource.HasFlipHorizontal)
                frame["flipH"] = JsonValue.Create(frameSource.FlipHorizontal);
            if (frameSource.HasFlipVertical)
                frame["flipV"] = JsonValue.Create(frameSource.FlipVertical);
            var projected = new JsonObject
            {
                ["id"] = StringNode(context.UniqueId(placeholder.Id)),
                ["name"] = StringNode(placeholder.Name),
                ["placeholderType"] = StringNode(PlaceholderType(placeholder.Type)),
                ["index"] = JsonValue.Create(placeholder.Index),
                ["frame"] = frame,
            };
            if (placeholder.Source is not null)
            {
                var capabilities = new List<CapabilitySpec>();
                // A direct placeholder transform is an owner-local leaf.  The
                // source binding has already rejected inherited/irregular
                // geometry, so changing the existing frame cannot silently
                // move a slide placeholder or rewrite its master graph.
                if (placeholder.Source.DirectFramePresenceEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type))
                    capabilities.Add(new("setFrame", EditableFrameFields));
                // Text is exposed only when the projection can preserve its
                // paragraph/run topology.  A string fallback for a formula,
                // hyperlink, or opaque run graph must remain source-owned.
                if (placeholder.Source.TextEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type) &&
                    PlaceholderTextProjectionEditable(placeholder.TextBody))
                    capabilities.Add(new("replaceText", ["text"]));
                if (placeholder.Source.TextEditable &&
                    BoundedLayoutPlaceholderType(placeholder.Type) &&
                    PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(placeholder.TextBody?.BodyProperties))
                    capabilities.Add(new("setTextBodyStyle", ["text.style"]));
                projected["nativeRef"] = NativeRef(
                    context,
                    $"placeholder:{placeholder.Id}",
                    HashOrFallback(placeholder.Source.ElementSha256, placeholder),
                    capabilities);
            }
            if (placeholder.TextBody is not null)
            {
                projected["text"] = TextContent(placeholder.TextBody, PptxTextCodec.Flatten(placeholder.TextBody), context);
                if (TextBoxStyle(placeholder.TextBody) is { Count: > 0 } style)
                    projected["style"] = style;
            }
            output.Add(projected);
        }
        return output;
    }

    private static bool PlaceholderTextProjectionEditable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase is
                PresentationTextRun.ContentOneofCase.Text or
                PresentationTextRun.ContentOneofCase.LineBreak or
                PresentationTextRun.ContentOneofCase.Field));

    private static bool BoundedLayoutPlaceholderType(string value) => value is
        "title" or "body" or "ctrTitle" or "subtitle" or "subTitle" or
        "content" or "obj" or "picture" or "pic" or "chart" or
        "table" or "tbl" or "date" or "dt" or "footer" or "ftr" or
        "slide-number" or "sldNum";

    private static void AddMasterTextLevels(
        JsonObject target,
        string name,
        IEnumerable<PresentationTextParagraph>? levels,
        ProjectionContext context)
    {
        if (levels is null) return;
        var projected = new JsonArray();
        foreach (var level in levels)
        {
            var item = new JsonObject();
            if (level.HasLevel) item["level"] = JsonValue.Create(checked((int)level.Level));
            if (level.HasAlignment && ParagraphAlignment(level.Alignment) is { } alignment)
                item["alignment"] = StringNode(alignment);
            if (level.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                item["indent"] = JsonValue.Create(Points(level.MarginLeftEmu));
            if (level.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                item["hanging"] = JsonValue.Create(-Points(level.IndentEmu));
            if (level.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                item["lineSpacing"] = JsonValue.Create(Math.Max(0.001, level.LineSpacingPoints));
            if (level.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                item["lineSpacingMultiplier"] = JsonValue.Create(Math.Max(0.00001, level.LineSpacingMultiplier));
            if (level.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                item["spaceBefore"] = JsonValue.Create(Math.Max(0, level.SpaceBeforePoints));
            if (level.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                item["spaceBeforeMultiplier"] = JsonValue.Create(Math.Max(0, level.SpaceBeforeMultiplier));
            if (level.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                item["spaceAfter"] = JsonValue.Create(Math.Max(0, level.SpaceAfterPoints));
            if (level.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                item["spaceAfterMultiplier"] = JsonValue.Create(Math.Max(0, level.SpaceAfterMultiplier));
            if (level.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                ProjectTextStyle(level.DefaultRunProperties) is { Count: > 0 } defaultText)
                item["defaultText"] = defaultText;
            if (ProjectBullet(level, context) is { } bullet) item["bullet"] = bullet;
            projected.Add(item);
        }
        if (projected.Count > 0) target[name] = projected;
    }

    private static JsonObject FrameDimensions(PresentationArtifact presentation) => new()
    {
        ["width"] = JsonValue.Create(Points(presentation.SlideWidthEmu)),
        ["height"] = JsonValue.Create(Points(presentation.SlideHeightEmu)),
        ["unit"] = StringNode("pt"),
    };

    private static JsonObject ProjectCanvas(PresentationArtifact presentation, ProjectionContext context)
    {
        var canvas = FrameDimensions(presentation);
        var objectHash = Sha256(CanonicalBytes(canvas));
        canvas["nativeRef"] = NativeRef(
            context,
            "canvas",
            objectHash,
            [new("setCanvas", ["canvas.width", "canvas.height"])]);
        return canvas;
    }

    private static void RegisterIds(PresentationArtifact presentation, ProjectionContext context)
    {
        foreach (var master in presentation.Masters) context.RegisterMaster(master.Id);
        foreach (var layout in presentation.Layouts) context.RegisterLayout(layout.Id);
        foreach (var customShow in presentation.CustomShows) context.RegisterCustomShow(customShow.Id);
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.RegisterPage(slide.Id, slide.Source?.PartPath);
            RegisterElementIds(slide.Elements, pageId, context);
        }
    }

    private static void RegisterElementIds(IEnumerable<PresentationElement> elements, string pageId, ProjectionContext context)
    {
        foreach (var element in elements)
        {
            context.RegisterElement(pageId, element.Id);
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                RegisterElementIds(element.Group.Children, pageId, context);
        }
    }

    private static JsonObject ProjectPage(
        PresentationSlide slide,
        PresentationArtifact presentation,
        ProjectionContext context)
    {
        var pageId = context.PageId(slide.Id);
        var pageHash = HashOrFallback(slide.Source?.SlideXmlSha256, slide);
        var pageCapabilities = new List<CapabilitySpec>();
        if (slide.Source is not null)
            pageCapabilities.Add(new("setName", ["name"]));
        if (slide.Source?.VisibilityEditable == true)
            pageCapabilities.Add(new("setHidden", ["hidden"]));
        if (slide.Source?.DeletionCapability?.Supported == true)
            pageCapabilities.Add(new("delete", ["element"]));
        if (slide.Source?.BackgroundEditable == true)
            pageCapabilities.Add(new("setBackground", ["background"]));
        if (slide.Source?.TransitionEditable == true || slide.Source?.TransitionAddable == true)
            pageCapabilities.Add(new("setTransition", ["transition"]));
        if (slide.Morph is null &&
            (slide.Source?.TimingEditable == true || slide.Source?.TimingAddable == true))
            pageCapabilities.Add(new("setAnimations", ["animations"]));
        if (slide.SpeakerNotes?.Source?.Editable == true || slide.Source?.SpeakerNotesAddable == true)
            pageCapabilities.Add(new("setNotes", ["notes"]));
        if (slide.Source is not null && presentation.Slides.Count > 1 && !presentation.SectionsOpaque)
            pageCapabilities.Add(new("reorder", ["pageOrder"]));
        if (slide.Source?.CloneCapability?.Supported == true)
            pageCapabilities.Add(new("duplicate", ["pageClone"]));
        if (slide.Source is not null)
            pageCapabilities.Add(new("appendElement", ["elements"]));

        var elements = new JsonArray();
        for (var elementIndex = 0; elementIndex < slide.Elements.Count; elementIndex++)
        {
            var element = slide.Elements[elementIndex];
            elements.Add(ProjectElement(
                element,
                slide,
                pageId,
                context,
                [element.Source?.ShapeTreeIndex ?? checked((uint)elementIndex)],
                element.ContentCase == PresentationElement.ContentOneofCase.Shape
                    ? EffectiveSlidePlaceholderFrame(element.Shape, slide, presentation)
                    : null));
        }

        var page = new JsonObject
        {
            ["id"] = StringNode(pageId),
            ["role"] = StringNode("source continuation"),
            ["elements"] = elements,
            ["nativeRef"] = NativeRef(context, $"page:{pageId}", pageHash, pageCapabilities),
        };
        if (slide.ReadingOrder.Count > 0)
        {
            var readingOrder = new JsonArray();
            foreach (var elementId in slide.ReadingOrder)
                if (context.TryElementId(pageId, elementId, out var projectedId))
                    readingOrder.Add(StringNode(projectedId));
            if (readingOrder.Count == slide.Elements.Count)
                page["readingOrder"] = readingOrder;
        }
        if (!string.IsNullOrWhiteSpace(slide.Name)) page["name"] = StringNode(slide.Name);
        if (!string.IsNullOrWhiteSpace(slide.LayoutId) && context.TryLayoutId(slide.LayoutId, out var layoutId))
            page["layout"] = StringNode(layoutId);
        if (slide.HasHidden) page["hidden"] = JsonValue.Create(slide.Hidden);
        if (!string.IsNullOrEmpty(slide.SpeakerNotes?.Text))
            page["notes"] = TextContent(slide.SpeakerNotes.TextBody, slide.SpeakerNotes.Text, context);
        if (ProjectBackground(slide.Background, context) is { } background) page["background"] = background;

        var animations = ProjectAnimations(slide, context);
        if (animations.Count > 0) page["animations"] = animations;
        if (ProjectTransition(slide, presentation, context) is { } transition) page["transition"] = transition;
        return page;
    }

    // Slide placeholders without a local a:xfrm render through the linked
    // layout, which in turn may inherit a matching master placeholder.  Keep
    // that effective frame in the PPJ view so the required public frame is
    // useful for inspection and editing.  The returned frame is only a
    // projection; the slide owner remains the source-bound edit boundary and
    // materialization is gated by the placeholder capability below.
    private static PresentationPlaceholderFrame? EffectiveSlidePlaceholderFrame(
        PresentationShape shape,
        PresentationSlide slide,
        PresentationArtifact presentation)
    {
        if (shape.Placeholder is null) return null;
        if (shape.DirectFrame is not null) return shape.DirectFrame.Clone();
        if (!shape.Placeholder.InheritsGeometry || string.IsNullOrWhiteSpace(slide.LayoutId)) return null;

        var layout = presentation.Layouts.FirstOrDefault(candidate =>
            candidate.Id.Equals(slide.LayoutId, StringComparison.Ordinal));
        if (layout is null) return null;
        var keyType = shape.Placeholder.Type;
        var keyIndex = shape.Placeholder.Index;
        var layoutMatches = layout.Placeholders
            .Where(candidate => candidate.Type.Equals(keyType, StringComparison.Ordinal) && candidate.Index == keyIndex)
            .ToArray();
        if (layoutMatches.Length > 1) return null;
        if (layoutMatches.Length == 1 && layoutMatches[0].DirectFrame is { } layoutFrame)
            return layoutFrame.Clone();

        var master = presentation.Masters.FirstOrDefault(candidate =>
            candidate.Id.Equals(layout.MasterId, StringComparison.Ordinal));
        if (master is null) return null;
        var masterMatches = master.Placeholders
            .Where(candidate => candidate.Type.Equals(keyType, StringComparison.Ordinal) && candidate.Index == keyIndex)
            .ToArray();
        return masterMatches.Length == 1 && masterMatches[0].DirectFrame is { } masterFrame
            ? masterFrame.Clone()
            : null;
    }

    private static JsonObject ProjectElement(
        PresentationElement element,
        PresentationSlide slide,
        string pageId,
        ProjectionContext context,
        IReadOnlyList<uint> shapeTreePath,
        PresentationPlaceholderFrame? effectivePlaceholderFrame = null)
    {
        var id = context.ElementId(pageId, element.Id);
        var hash = HashOrFallback(element.Source?.ElementSha256, element);
        var customSites = element.Connector is { } connector &&
            (connector.StartTargetId.Length != 0 || connector.EndTargetId.Length != 0) &&
            (connector.StartTargetId.Length == 0 || ProjectableCustomSite(slide, connector.StartTargetId, connector.StartConnectionSiteIndex)) &&
            (connector.EndTargetId.Length == 0 || ProjectableCustomSite(slide, connector.EndTargetId, connector.EndConnectionSiteIndex));
        var capabilities = Capabilities(element, effectivePlaceholderFrame is not null, customSites);
        var leaves = PpjNativeLeafProjection.Describe(
            context.SourceSha256,
            pageId,
            id,
            element,
            shapeTreePath,
            context.RecordNativeLeaf);
        var nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, capabilities, leaves);
        JsonObject projected = element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => ProjectShape(element, id, nativeRef, context, effectivePlaceholderFrame),
            PresentationElement.ContentOneofCase.Image => ProjectImage(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Table => ProjectTable(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Connector => ProjectConnector(element, id, nativeRef, pageId, context, slide),
            PresentationElement.ContentOneofCase.Chart => ProjectChart(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Diagram => ProjectNativeSmartArt(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Group => ProjectGroup(element, id, nativeRef, slide, pageId, context, shapeTreePath),
            PresentationElement.ContentOneofCase.Opaque when element.Opaque.DiagramText is not null =>
                ProjectSourceSmartArt(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Opaque when element.Opaque.OleWorkbook is not null || element.Opaque.OleOfficePackage is not null =>
                ProjectSourceOle(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Opaque => ProjectOpaque(element, id, nativeRef),
            _ => ProjectOpaque(element, id, nativeRef, "unknown"),
        };
        if (element.HasHidden && element.Hidden) projected["hidden"] = JsonValue.Create(true);
        if (element.HasLocked && element.Locked) projected["locked"] = JsonValue.Create(true);
        if (projected["type"]!.GetValue<string>() == "opaque")
        {
            // ProjectShape/ProjectImage may conservatively fall back to opaque
            // after the first nativeRef already owns this JsonArray. Clone the
            // issued leaf descriptors; a JsonNode cannot have two parents.
            nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, OpaqueCapabilities(element), leaves.DeepClone().AsArray());
            projected["nativeRef"] = nativeRef;
        }
        context.RecordNode(pageId, id, projected["type"]!.GetValue<string>(), nativeRef);
        return projected;
    }

    private static JsonObject ProjectShape(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context,
        PresentationPlaceholderFrame? effectivePlaceholderFrame = null)
    {
        var shape = element.Shape;
        var isPlaceholder = shape.Placeholder is not null;
        if (isPlaceholder && shape.DirectFrame is null && effectivePlaceholderFrame is null)
            return ProjectOpaque(element, id, nativeRef, "placeholder", "Preserved slide placeholder whose effective layout/master frame is ambiguous or unavailable.");
        var frame = isPlaceholder && (effectivePlaceholderFrame ?? shape.DirectFrame) is { } placeholderFrame
            ? ShapeFrame(placeholderFrame)
            : ShapeFrame(shape);
        var text = TextContent(shape.TextBody, shape.Text, context);
        var hasText = !string.IsNullOrEmpty(shape.Text) || shape.TextBody?.Paragraphs.Count > 0;
        var isTextBox = shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry);
        var lineLike = shape.Geometry == "line" || PpjLinePathCodec.IsLineLike(shape);
        // The importer deliberately leaves unsupported custom path graphs
        // empty.  The source-bound projection can still expose the owning
        // shape (frame, native leaves, and safe paint fields) without
        // pretending that its path topology is editable.
        var sourceCustomGeometry = shape.Geometry == "custom";
        var sourceImageFill = !string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
            context.TryMaterializeAsset(shape.ImageFillAssetId, out var sourceImageAssetId);
        if (shape.ImageFill is not null && !context.TryMaterializeAsset(shape.ImageFill.AssetId, out _))
            return ProjectOpaque(element, id, nativeRef, "shape", "Preserved source shape whose image fill cannot be materialized safely.");
        if (!isPlaceholder && !isTextBox && !PptxPresetGeometryAdjustmentCodec.HasProfile(shape.Geometry) &&
            !CanProjectCustomGeometry(shape, allowShapeGraph: true) && !sourceImageFill && !sourceCustomGeometry && !lineLike)
            return ProjectOpaque(element, id, nativeRef, "shape", $"Preserved source shape with unsupported geometry '{shape.Geometry}'.");

        var common = ElementBase(id, element.Name, frame, Accessibility(shape.Accessibility), nativeRef);
        if (ProjectAction(shape.Action, context) is { } action)
            common["action"] = action;
        if (ProjectAction(shape.HoverAction, context) is { } hoverAction)
            common["hoverAction"] = hoverAction;
        if (shape.Placeholder is not null)
        {
            common["type"] = StringNode("placeholder");
            common["placeholderType"] = StringNode(PlaceholderType(shape.Placeholder.Type));
            common["index"] = JsonValue.Create(shape.Placeholder.Index);
            if (hasText) common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } placeholderStyle) common["style"] = placeholderStyle;
            return common;
        }
        if (shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry))
        {
            common["type"] = StringNode("text");
            common["text"] = text;
            if (TextBoxStyle(shape.TextBody) is { Count: > 0 } textStyle) common["style"] = textStyle;
            ApplyTextContainerStyle(common, shape, context);
            return common;
        }
        if (lineLike)
        {
            common["type"] = StringNode("line");
            if (shape.Geometry == "line")
                common["path"] = PpjLinePathCodec.Synthetic(shape);
            else if (PpjLinePathCodec.TryProjectKimi(shape) is { } kimiPath)
            {
                common["viewBox"] = kimiPath["viewBox"]!.DeepClone();
                common["points"] = kimiPath["points"]!.DeepClone();
                common["curve"] = kimiPath["curve"]!.DeepClone();
            }
            else
                common["path"] = PpjLinePathCodec.Project(shape);
            var lineStyle = ShapeStyle(shape, context);
            if (lineStyle.TryGetPropertyValue("stroke", out var lineStroke) && lineStroke is not null)
                common["stroke"] = lineStroke.DeepClone();
            else
                common["stroke"] = Stroke("000000", shape.LineWidthEmu, "solid", shape.LineCap, shape.LineJoin,
                    shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                    shape.LineScheme);
            if (Arrow(shape.StartArrow) is { } startArrow) common["startArrow"] = startArrow;
            if (Arrow(shape.EndArrow) is { } endArrow) common["endArrow"] = endArrow;
            if (shape.Shadow is not null) common["shadow"] = Shadow(shape.Shadow);
            if (shape.Glow is not null) common["glow"] = Glow(shape.Glow);
            if (shape.InnerShadow is not null) common["innerShadow"] = InnerShadow(shape.InnerShadow);
            if (shape.Reflection is not null) common["reflection"] = Reflection(shape.Reflection);
            if (shape.SoftEdge is not null) common["softEdge"] = SoftEdge(shape.SoftEdge);
            return common;
        }
        common["type"] = StringNode("shape");
        var compoundOpacity = 1d;
        var hasCompoundOpacity = element.Source?.Editable == true &&
            TryGetCompoundShapeOpacity(shape, out compoundOpacity);
        if (shape.Geometry == "custom")
            common["geometry"] = CanProjectCustomGeometry(shape, allowShapeGraph: true)
                ? ProjectCustomGeometry(shape)
                : new JsonObject
                {
                    // The path graph is intentionally not guessed when a
                    // source-bound image-filled custom shape falls outside
                    // the bounded custom-geometry profile. Keep the shape's
                    // native geometry behind an explicit source-bound marker
                    // while exposing its frame, image asset and capabilities.
                    ["kind"] = StringNode("source-custom"),
                    ["sourceBound"] = JsonValue.Create(true),
                };
        else
        {
            var geometry = new JsonObject { ["kind"] = StringNode("preset"), ["preset"] = StringNode(shape.Geometry) };
            if (PptxPresetGeometryAdjustmentCodec.IsCompleteValues(shape.Geometry, shape.PresetAdjustments))
                geometry["adjustments"] = new JsonArray(shape.PresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            common["geometry"] = geometry;
        }
        if (hasText) common["text"] = text;
        if (TextBoxStyle(shape.TextBody) is { Count: > 0 } shapeTextStyle) common["textStyle"] = shapeTextStyle;
        var style = ShapeStyle(shape, context, omitOwnerOpacity: hasCompoundOpacity);
        if (style.Count > 0) common["style"] = style;
        if (hasCompoundOpacity)
            common["compositing"] = new JsonObject { ["opacity"] = JsonValue.Create(compoundOpacity) };
        return common;
    }

    internal static bool TryGetCompoundShapeOpacity(PresentationShape shape, out double opacity)
    {
        opacity = 1;
        if (shape.Placeholder is not null || !string.IsNullOrEmpty(shape.Text) ||
            shape.Geometry is "line" or "custom" or "textbox" or "none" ||
            string.IsNullOrEmpty(shape.Geometry))
            return false;
        if (!string.IsNullOrEmpty(shape.FillRgb) && !string.IsNullOrEmpty(shape.FillScheme))
            return false;
        if (!string.IsNullOrWhiteSpace(shape.ImageFillAssetId) && shape.ImageFill is null)
            return false;

        var values = new List<double>();
        if (shape.FillRgb.Length > 0 || shape.FillScheme.Length > 0)
            values.Add(shape.HasFillOpacityThousandthPercent ? Unit(shape.FillOpacityThousandthPercent) : 1);
        if (shape.GradientFill is not null)
        {
            if (shape.GradientFill.Stops.Count == 0) return false;
            values.AddRange(shape.GradientFill.Stops.Select(stop =>
                stop.HasOpacityThousandthPercent ? Unit(stop.OpacityThousandthPercent) : 1));
        }
        if (shape.ImageFill is not null)
            values.Add(shape.ImageFill.HasOpacityThousandthPercent ? Unit(shape.ImageFill.OpacityThousandthPercent) : 1);
        if (shape.LineStyle != "none" && (shape.LineRgb.Length > 0 || shape.LineScheme.Length > 0))
            values.Add(shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : 1);
        if (shape.Shadow is not null)
            values.Add(shape.Shadow.HasOpacityThousandthPercent ? Unit(shape.Shadow.OpacityThousandthPercent) : 1);
        if (values.Count == 0) return false;
        var candidate = values[0];
        opacity = candidate;
        return values.All(value => Math.Abs(value - candidate) < 0.000005);
    }

    private static bool CanProjectCustomGeometry(PresentationShape shape, bool allowShapeGraph = false)
    {
        if (shape.Geometry != "custom" || shape.CustomPaths.Count == 0)
            return false;
        var width = shape.CustomPaths[0].Width;
        var height = shape.CustomPaths[0].Height;
        if (!allowShapeGraph && (width <= 0 || height <= 0 || width > 100_000_000 || height > 100_000_000 || shape.CustomPaths.Any(path => path.Width != width || path.Height != height)))
            return false;
        return shape.CustomPaths.SelectMany(path => path.Commands).All(command => command.CommandCase switch
        {
            PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => allowShapeGraph || Literal(command.MoveTo),
            PresentationCustomGeometryCommand.CommandOneofCase.LineTo => allowShapeGraph || Literal(command.LineTo),
            PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo =>
                allowShapeGraph || Literal(command.QuadraticBezierTo.Control) && Literal(command.QuadraticBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo =>
                allowShapeGraph || Literal(command.CubicBezierTo.Control1) && Literal(command.CubicBezierTo.Control2) && Literal(command.CubicBezierTo.End),
            PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => allowShapeGraph || Literal(command.ArcTo),
            PresentationCustomGeometryCommand.CommandOneofCase.Close => true,
            _ => false,
        });
    }

    private static bool Literal(PresentationCustomGeometryPoint point) =>
        !point.HasXReference && !point.HasYReference;

    private static bool Literal(PresentationCustomGeometryArc arc) =>
        !arc.HasWidthRadiusReference && !arc.HasHeightRadiusReference &&
        !arc.HasStartAngleReference && !arc.HasSweepAngleReference;

    private static double ProjectPathBaseline(long extent) =>
        extent is > 0 and <= 100_000_000 ? extent / 1_000d : 1d;

    private static JsonObject ProjectCustomGeometry(PresentationShape shape)
    {
        var width = ProjectPathBaseline(shape.CustomPaths[0].Width);
        var height = ProjectPathBaseline(shape.CustomPaths[0].Height);
        var paths = new JsonArray();
        foreach (var source in shape.CustomPaths)
        {
            var commands = new JsonArray();
            foreach (var command in source.Commands) commands.Add(ProjectCustomCommand(command));
            var path = new JsonObject { ["commands"] = commands };
            if (source.Width / 1_000d != width || source.Height / 1_000d != height)
                path["viewport"] = new JsonObject
                {
                    ["width"] = source.Width / 1_000d, ["height"] = source.Height / 1_000d,
                };
            if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.Normal) path["fill"] = JsonValue.Create(true);
            else if (source.FillMode == PresentationCustomGeometryPath.Types.FillMode.None) path["fill"] = JsonValue.Create(false);
            if (source.HasStroke) path["stroke"] = JsonValue.Create(source.Stroke);
            if (source.HasExtrusionAllowed) path["extrusionOk"] = JsonValue.Create(source.ExtrusionAllowed);
            paths.Add(path);
        }
        var output = new JsonObject
        {
            ["kind"] = StringNode("custom"),
            ["viewBox"] = new JsonObject
            {
                ["x"] = JsonValue.Create(0),
                ["y"] = JsonValue.Create(0),
                ["width"] = JsonValue.Create(width),
                ["height"] = JsonValue.Create(height),
            },
            ["paths"] = paths,
        };
        if (shape.CustomAdjustmentHandles.Count > 0)
            output["adjustmentHandles"] = new JsonArray(shape.CustomAdjustmentHandles.Select(handle =>
                (JsonNode)PpjCustomGeometryHandleCodec.Project(handle)).ToArray());
        if (shape.CustomConnectionSites.Count > 0)
            output["connectionSites"] = new JsonArray(shape.CustomConnectionSites.Select(site => (JsonNode)new JsonObject
            {
                ["angle"] = site.HasAngleReference ? JsonValue.Create(site.AngleReference) : JsonValue.Create(site.Angle60000 / 60_000d),
                ["x"] = site.HasXReference ? JsonValue.Create(site.XReference) : JsonValue.Create(site.XEmu / 12_700d),
                ["y"] = site.HasYReference ? JsonValue.Create(site.YReference) : JsonValue.Create(site.YEmu / 12_700d),
            }).ToArray());
        if (shape.CustomAdjustments.Count > 0)
            output["adjustments"] = new JsonArray(shape.CustomAdjustments.Select(adjustment => (JsonNode)new JsonObject
            {
                ["name"] = adjustment.Name, ["formula"] = adjustment.Formula,
            }).ToArray());
        if (shape.CustomGuides.Count > 0)
            output["guides"] = new JsonArray(shape.CustomGuides.Select(guide => (JsonNode)new JsonObject
            {
                ["name"] = guide.Name, ["formula"] = guide.Formula,
            }).ToArray());
        if (shape.TextRectangle is { } rectangle)
            output["textRectangle"] = new JsonObject
            {
                ["left"] = rectangle.HasLeftReference ? JsonValue.Create(rectangle.LeftReference) : JsonValue.Create(rectangle.LeftEmu / 12_700d),
                ["top"] = rectangle.HasTopReference ? JsonValue.Create(rectangle.TopReference) : JsonValue.Create(rectangle.TopEmu / 12_700d),
                ["right"] = rectangle.HasRightReference ? JsonValue.Create(rectangle.RightReference) : JsonValue.Create(rectangle.RightEmu / 12_700d),
                ["bottom"] = rectangle.HasBottomReference ? JsonValue.Create(rectangle.BottomReference) : JsonValue.Create(rectangle.BottomEmu / 12_700d),
            };
        return output;
    }

    private static JsonObject ProjectCustomCommand(PresentationCustomGeometryCommand command) => command.CommandCase switch
    {
        PresentationCustomGeometryCommand.CommandOneofCase.MoveTo => ProjectCustomPoint("moveTo", command.MoveTo),
        PresentationCustomGeometryCommand.CommandOneofCase.LineTo => ProjectCustomPoint("lineTo", command.LineTo),
        PresentationCustomGeometryCommand.CommandOneofCase.QuadraticBezierTo => new JsonObject
        {
            ["op"] = StringNode("quadraticTo"),
            ["x1"] = CustomPathValue(command.QuadraticBezierTo.Control.HasXReference, command.QuadraticBezierTo.Control.XReference, CustomPathPoint(command.QuadraticBezierTo.Control.X)),
            ["y1"] = CustomPathValue(command.QuadraticBezierTo.Control.HasYReference, command.QuadraticBezierTo.Control.YReference, CustomPathPoint(command.QuadraticBezierTo.Control.Y)),
            ["x"] = CustomPathValue(command.QuadraticBezierTo.End.HasXReference, command.QuadraticBezierTo.End.XReference, CustomPathPoint(command.QuadraticBezierTo.End.X)),
            ["y"] = CustomPathValue(command.QuadraticBezierTo.End.HasYReference, command.QuadraticBezierTo.End.YReference, CustomPathPoint(command.QuadraticBezierTo.End.Y)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.CubicBezierTo => new JsonObject
        {
            ["op"] = StringNode("cubicTo"),
            ["x1"] = CustomPathValue(command.CubicBezierTo.Control1.HasXReference, command.CubicBezierTo.Control1.XReference, CustomPathPoint(command.CubicBezierTo.Control1.X)),
            ["y1"] = CustomPathValue(command.CubicBezierTo.Control1.HasYReference, command.CubicBezierTo.Control1.YReference, CustomPathPoint(command.CubicBezierTo.Control1.Y)),
            ["x2"] = CustomPathValue(command.CubicBezierTo.Control2.HasXReference, command.CubicBezierTo.Control2.XReference, CustomPathPoint(command.CubicBezierTo.Control2.X)),
            ["y2"] = CustomPathValue(command.CubicBezierTo.Control2.HasYReference, command.CubicBezierTo.Control2.YReference, CustomPathPoint(command.CubicBezierTo.Control2.Y)),
            ["x"] = CustomPathValue(command.CubicBezierTo.End.HasXReference, command.CubicBezierTo.End.XReference, CustomPathPoint(command.CubicBezierTo.End.X)),
            ["y"] = CustomPathValue(command.CubicBezierTo.End.HasYReference, command.CubicBezierTo.End.YReference, CustomPathPoint(command.CubicBezierTo.End.Y)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.ArcTo => new JsonObject
        {
            ["op"] = StringNode("arcTo"),
            ["radiusX"] = CustomPathValue(command.ArcTo.HasWidthRadiusReference, command.ArcTo.WidthRadiusReference, CustomPathPoint(command.ArcTo.WidthRadius)),
            ["radiusY"] = CustomPathValue(command.ArcTo.HasHeightRadiusReference, command.ArcTo.HeightRadiusReference, CustomPathPoint(command.ArcTo.HeightRadius)),
            ["startAngle"] = CustomPathValue(command.ArcTo.HasStartAngleReference, command.ArcTo.StartAngleReference, NormalizeCustomPathStartAngle(command.ArcTo.StartAngle)),
            ["sweepAngle"] = CustomPathValue(command.ArcTo.HasSweepAngleReference, command.ArcTo.SweepAngleReference, CustomPathAngle(command.ArcTo.SweepAngle)),
        },
        PresentationCustomGeometryCommand.CommandOneofCase.Close => new JsonObject { ["op"] = StringNode("close") },
        _ => throw new InvalidOperationException("Unsupported PPJ custom path command passed the projection gate."),
    };

    private static JsonObject ProjectCustomPoint(string operation, PresentationCustomGeometryPoint point) => new()
    {
        ["op"] = StringNode(operation),
        ["x"] = CustomPathValue(point.HasXReference, point.XReference, CustomPathPoint(point.X)),
        ["y"] = CustomPathValue(point.HasYReference, point.YReference, CustomPathPoint(point.Y)),
    };

    private static JsonNode CustomPathValue(bool hasReference, string reference, double literal) =>
        hasReference ? JsonValue.Create(reference)! : JsonValue.Create(literal)!;

    private static double CustomPathPoint(long value) => value / 1_000d;

    private static double CustomPathAngle(int value) => value / 60_000d;

    private static double NormalizeCustomPathStartAngle(int value)
    {
        var degrees = CustomPathAngle(value);
        return ((degrees % 360) + 360) % 360;
    }

    private static JsonObject ProjectImage(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var image = element.Image;
        if (!context.TryMaterializeAsset(image.AssetId, out var assetId))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose media payload cannot be materialized safely.");
        string? svgAssetId = null;
        if (!string.IsNullOrEmpty(image.SvgAssetId) && !context.TryMaterializeAsset(image.SvgAssetId, out svgAssetId))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose paired SVG payload cannot be materialized safely.");
        var customMask = image.CustomMaskPaths.Count > 0 ? ImageMaskShape(image) : null;
        if (customMask is not null && !CanProjectCustomGeometry(customMask))
            return ProjectOpaque(element, id, nativeRef, "picture", "Preserved source picture whose custom mask cannot be represented exactly in PPJ.");
        var output = ElementBase(id, element.Name, ImageFrame(image), ImageAccessibility(image), nativeRef);
        output["type"] = StringNode("image");
        output["asset"] = StringNode(assetId);
        if (svgAssetId is not null) output["svgAsset"] = StringNode(svgAssetId);
        output["fit"] = StringNode(image.Tiled ? "tile" : "stretch");
        if (image.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(image.Crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(image.Crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(image.Crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(image.Crop.BottomThousandthPercent)),
            };
        }
        if (image.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(image.OpacityThousandthPercent));
        if (customMask is not null)
            output["mask"] = ProjectCustomGeometry(customMask);
        else if (!string.IsNullOrEmpty(image.MaskPreset) && PptxPresetGeometryAdjustmentCodec.HasProfile(image.MaskPreset))
        {
            var mask = new JsonObject { ["kind"] = StringNode("preset"), ["preset"] = StringNode(image.MaskPreset) };
            if (image.MaskPresetAdjustments.Count > 0)
                mask["adjustments"] = new JsonArray(image.MaskPresetAdjustments.Select(value => JsonValue.Create(value)).ToArray());
            output["mask"] = mask;
        }
        if (image.Border is not null &&
            (!string.IsNullOrEmpty(image.Border.ColorRgb) || !string.IsNullOrEmpty(image.Border.ColorScheme)))
            output["border"] = Stroke(image.Border.ColorRgb, image.Border.WidthEmu, image.Border.Style, image.Border.Cap, image.Border.Join,
                image.Border.HasOpacityThousandthPercent ? Unit(image.Border.OpacityThousandthPercent) : null,
                image.Border.ColorScheme);
        if (image.Shadow is not null && (!string.IsNullOrEmpty(image.Shadow.ColorRgb) || !string.IsNullOrEmpty(image.Shadow.ColorScheme)))
            output["shadow"] = Shadow(image.Shadow);
        if (image.Glow is not null && (!string.IsNullOrEmpty(image.Glow.ColorRgb) || !string.IsNullOrEmpty(image.Glow.ColorScheme)))
            output["glow"] = Glow(image.Glow);
        if (image.InnerShadow is not null && (!string.IsNullOrEmpty(image.InnerShadow.ColorRgb) || !string.IsNullOrEmpty(image.InnerShadow.ColorScheme)))
            output["innerShadow"] = InnerShadow(image.InnerShadow);
        if (image.Reflection is not null)
            output["reflection"] = Reflection(image.Reflection);
        if (image.SoftEdge is not null)
            output["softEdge"] = SoftEdge(image.SoftEdge);
        return output;
    }

    private static PresentationShape ImageMaskShape(PresentationImage image)
    {
        var shape = new PresentationShape
        {
            Geometry = "custom",
            WidthEmu = image.WidthEmu,
            HeightEmu = image.HeightEmu,
        };
        shape.CustomPaths.Add(image.CustomMaskPaths);
        return shape;
    }

    private static JsonObject ProjectChart(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var chart = element.Chart;
        var type = ChartType(chart.Type);
        var series = chart.Type == SpreadsheetChartType.Combo
            ? chart.ComboSeries.Select(item => item.Series).ToArray()
            : chart.Series.ToArray();
        var numericX = chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;
        var invalidSeries = numericX
            ? chart.Categories.Count != 0 || series.Any(item =>
                item.Values.Count != item.XValues.Count ||
                chart.Type == SpreadsheetChartType.Bubble && item.Values.Count != item.BubbleSizes.Count ||
                chart.Type == SpreadsheetChartType.Scatter && item.BubbleSizes.Count != 0)
            : series.Any(item => item.Values.Count != chart.Categories.Count || item.XValues.Count != 0 || item.BubbleSizes.Count != 0);
        if (type is null || series.Length == 0 || invalidSeries)
            return ProjectOpaque(element, id, nativeRef, "chart", "Preserved source chart outside the bounded PPJ data profile.");

        var output = ElementBase(id, element.Name, ChartFrame(chart), Accessibility(chart.Accessibility), nativeRef);
        output["type"] = StringNode("chart");
        output["chartType"] = StringNode(chart.Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar" ? "bar" : type!);
        if (!string.IsNullOrEmpty(chart.Title))
            output["title"] = chart.TitleBody is null
                ? StringNode(chart.Title)
                : TextContent(chart.TitleBody, chart.Title, context);
        output["titlePlacement"] = StringNode(chart.Title.Length == 0
            ? "none"
            : chart.HasTitlePlacement && chart.TitlePlacement.Length > 0
                ? chart.TitlePlacement
                : "aboveChart");
        if (chart.TextStyle is { FontFamily.Length: > 0 } globalTextStyle)
            output["fontFamily"] = StringNode(globalTextStyle.FontFamily);
        var projectSeriesNullHandling = (chart.Type is SpreadsheetChartType.Line or SpreadsheetChartType.Area or SpreadsheetChartType.Radar) &&
            chart.HasDisplayBlanksAs && chart.DisplayBlanksAs is ("zero" or "gap" or "span");
        if (chart.HasDisplayBlanksAs && !projectSeriesNullHandling)
            output["displayBlanksAs"] = StringNode(chart.DisplayBlanksAs);
        if (chart.HasStyleIndex) output["styleIndex"] = JsonValue.Create(chart.StyleIndex);
        if (chart.HasRoundedCorners) output["roundedCorners"] = JsonValue.Create(chart.RoundedCorners);
        if (chart.DataTable is not null) output["dataTable"] = ProjectChartDataTable(chart.DataTable);
        var categories = new JsonArray();
        foreach (var value in chart.Categories) categories.Add(StringNode(value));
        var seriesJson = new JsonArray();
        for (var index = 0; index < series.Length; index++)
        {
            var item = series[index];
            var values = new JsonArray();
            var missingValueIndexes = item.MissingValueIndexes.ToHashSet();
            for (var valueIndex = 0; valueIndex < item.Values.Count; valueIndex++)
                values.Add(missingValueIndexes.Contains((uint)valueIndex) ? null : NumberNode(item.Values[valueIndex]));
            var entry = new JsonObject
            {
                ["id"] = StringNode($"series-{index + 1}"),
                ["name"] = StringNode(item.Name ?? string.Empty),
                ["values"] = values,
            };
            if (projectSeriesNullHandling)
                entry["nullHandling"] = StringNode(chart.DisplayBlanksAs == "span" ? "connect" : chart.DisplayBlanksAs);
            if (!string.IsNullOrWhiteSpace(item.CategoryFormula)) entry["categoryFormula"] = StringNode(item.CategoryFormula);
            if (!string.IsNullOrWhiteSpace(item.XValueFormula)) entry["xValueFormula"] = StringNode(item.XValueFormula);
            if (!string.IsNullOrWhiteSpace(item.ValueFormula)) entry["valueFormula"] = StringNode(item.ValueFormula);
            if (!string.IsNullOrWhiteSpace(item.BubbleSizeFormula)) entry["bubbleSizeFormula"] = StringNode(item.BubbleSizeFormula);
            if (item.XValues.Count > 0)
            {
                var xValues = new JsonArray();
                foreach (var value in item.XValues) xValues.Add(NumberNode(value));
                entry["xValues"] = xValues;
            }
            if (item.BubbleSizes.Count > 0)
            {
                var bubbleSizes = new JsonArray();
                foreach (var value in item.BubbleSizes) bubbleSizes.Add(NumberNode(value));
                entry["bubbleSizes"] = bubbleSizes;
            }
            if (chart.Type == SpreadsheetChartType.Combo)
            {
                entry["chartType"] = StringNode(chart.ComboSeries[index].Type == SpreadsheetChartType.Bar && chart.BarDirection == "bar"
                    ? "bar"
                    : ChartType(chart.ComboSeries[index].Type) ?? "line");
                var axisIndex = chart.ComboSeries[index].AxisGroup == PresentationChartAxisGroup.Secondary ? 1 : 0;
                entry["axis"] = StringNode(axisIndex == 1 ? "secondary" : "primary");
                // The native chart model has two axis groups rather than the
                // Kimi array form. Emit both bounded indexes so a projected
                // combo series can be fed back into the dataset/encode
                // compatibility layer without losing its 0/1 placement.
                entry["xAxisIndex"] = JsonValue.Create(axisIndex);
                entry["yAxisIndex"] = JsonValue.Create(axisIndex);
            }
            ProjectChartSeriesStyle(entry, item);
            seriesJson.Add(entry);
        }
        var data = new JsonObject { ["categories"] = categories, ["series"] = seriesJson };
        if (CanProjectCanonicalDataset(chart, series))
        {
            // Keep the legacy categories/series projection for wire
            // compatibility, and add a deterministic dataset/encoding view
            // only when it is a lossless view of the same native vectors.
            // Complex formulas, sparse point overrides, and per-series
            // topology stay on the legacy path so a re-import cannot silently
            // discard their semantics.
            data["dataset"] = CanonicalDataset(chart, series);
            var encoding = chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble
                ? new JsonObject
                {
                    ["x"] = StringNode("x"),
                    ["y"] = StringNode("y"),
                    ["series"] = StringNode("series"),
                }
                : new JsonObject
                {
                    ["category"] = StringNode("category"),
                    ["series"] = StringNode("series"),
                    ["value"] = StringNode("value"),
                };
            if (chart.Type == SpreadsheetChartType.Bubble) encoding["size"] = StringNode("size");
            data["encoding"] = encoding;
        }
        output["data"] = data;
        if (chart.Type == SpreadsheetChartType.Radar && TryProjectRadarSpokeAxis(chart, out var spokeAxis))
            output["spokeAxis"] = spokeAxis;
        else
        {
            if (chart.XAxis is not null) output["xAxis"] = ProjectChartAxis(chart.XAxis);
            if (chart.YAxis is not null) output["yAxis"] = ProjectChartAxis(chart.YAxis);
        }
        if (chart.SecondaryXAxis is not null) output["secondaryXAxis"] = ProjectChartAxis(chart.SecondaryXAxis);
        if (chart.SecondaryYAxis is not null) output["secondaryYAxis"] = ProjectChartAxis(chart.SecondaryYAxis);
        var style = new JsonObject
        {
            ["legend"] = StringNode(chart.HasLegend
                ? chart.LegendPosition.Length == 0 ? "right" : chart.LegendPosition
                : "none"),
        };
        if (chart.HasLegendOverlay) style["legendOverlay"] = JsonValue.Create(chart.LegendOverlay);
        if (chart.LegendFill is not null) style["legendFill"] = ProjectChartSurfaceFill(chart.LegendFill);
        if (chart.LegendLine is not null && !string.IsNullOrEmpty(chart.LegendLine.Color?.Rgb))
            style["legendLine"] = ProjectChartLine(chart.LegendLine);
        if (chart.Grouping.Length > 0) style["stacking"] = StringNode(chart.Grouping);
        if (chart.HasGapWidth) style["gapWidth"] = JsonValue.Create(chart.GapWidth);
        if (chart.HasOverlap) style["overlap"] = JsonValue.Create(chart.Overlap);
        if (chart.HasVaryColors) style["varyColors"] = JsonValue.Create(chart.VaryColors);
        if (chart.HasFirstSliceAngle) style["startAngle"] = JsonValue.Create(chart.FirstSliceAngle);
        if (chart.HasDoughnutHoleSize) style["holeSize"] = JsonValue.Create(chart.DoughnutHoleSize);
        if (chart.HasBubbleScale) style["bubbleScale"] = JsonValue.Create(chart.BubbleScale);
        if (chart.BubbleSizeMode.Length > 0) style["bubbleSizeMode"] = StringNode(chart.BubbleSizeMode);
        if (chart.ScatterStyle.Length > 0) style["scatterStyle"] = StringNode(chart.ScatterStyle);
        if (chart.RadarStyle.Length > 0) style["radarStyle"] = StringNode(chart.RadarStyle);
        if (chart.XAxis is null && chart.HasShowCategoryAxis) style["showCategoryAxis"] = JsonValue.Create(chart.ShowCategoryAxis);
        if (chart.YAxis is null && chart.HasShowValueAxis) style["showValueAxis"] = JsonValue.Create(chart.ShowValueAxis);
        if (chart.YAxis is null && chart.HasShowGridlines) style["showGridlines"] = JsonValue.Create(chart.ShowGridlines);
        if (chart.ChartAreaFill is not null) style["chartAreaFill"] = ProjectChartSurfaceFill(chart.ChartAreaFill);
        if (chart.PlotAreaFill is not null) style["plotAreaFill"] = ProjectChartSurfaceFill(chart.PlotAreaFill);
        if (chart.PlotAreaLine is not null && !string.IsNullOrEmpty(chart.PlotAreaLine.Color?.Rgb))
            style["plotAreaLine"] = ProjectChartLine(chart.PlotAreaLine);
        if (chart.Frame is not null)
        {
            var frame = new JsonObject();
            if (chart.Frame.ImageFill is not null)
            {
                if (ProjectImagePaint(chart.Frame.ImageFill, context) is { } imageFill)
                    frame["fill"] = imageFill;
            }
            else if (chart.Frame.Fill is not null) frame["fill"] = ProjectChartSurfaceFill(chart.Frame.Fill);
            if (chart.Frame.Line is not null) frame["stroke"] = ProjectChartLine(chart.Frame.Line);
            if (chart.Frame.Shadow is not null) frame["shadow"] = Shadow(chart.Frame.Shadow);
            if (frame.Count > 0) style["frame"] = frame;
        }
        if (chart.TitleTextStyle is not null)
            style["titleTextStyle"] = ProjectChartTextStyle(chart.TitleTextStyle);
        if (chart.LegendTextStyle is not null)
            style["legendTextStyle"] = ProjectChartTextStyle(chart.LegendTextStyle);
        if (chart.LineOptions?.HasSmooth == true) style["smooth"] = JsonValue.Create(chart.LineOptions.Smooth);
        if (chart.LineOptions?.VaryColors == true) style["varyColors"] = JsonValue.Create(true);
        if (chart.DataLabels is not null)
        {
            var labels = new JsonObject
            {
                ["showValue"] = JsonValue.Create(chart.DataLabels.ShowValue),
                ["showCategory"] = JsonValue.Create(chart.DataLabels.ShowCategoryName),
            };
            if (chart.DataLabels.HasShowSeriesName) labels["showSeries"] = JsonValue.Create(chart.DataLabels.ShowSeriesName);
            if (chart.DataLabels.HasShowPercent) labels["showPercent"] = JsonValue.Create(chart.DataLabels.ShowPercent);
            if (chart.DataLabels.HasShowBubbleSize) labels["showBubbleSize"] = JsonValue.Create(chart.DataLabels.ShowBubbleSize);
            if (chart.DataLabels.HasShowLeaderLines) labels["showLeaderLines"] = JsonValue.Create(chart.DataLabels.ShowLeaderLines);
            if (chart.DataLabels.HasPosition && DataLabelPosition(chart.DataLabels.Position) is { } position)
                labels["position"] = StringNode(position);
            if (chart.DataLabels.TextStyle is not null)
                labels["textStyle"] = ProjectChartTextStyle(chart.DataLabels.TextStyle);
            if (chart.DataLabels.NumberFormatCode.Length > 0)
                labels["numberFormat"] = StringNode(chart.DataLabels.NumberFormatCode);
            if (chart.DataLabels.Fill is not null)
                labels["fill"] = ProjectChartSurfaceFill(chart.DataLabels.Fill);
            if (chart.DataLabels.Line is not null && !string.IsNullOrEmpty(chart.DataLabels.Line.Color?.Rgb))
                labels["line"] = ProjectChartLine(chart.DataLabels.Line);
            style["dataLabels"] = labels;
        }
        output["style"] = style;
        return output;
    }

    private static bool CanProjectCanonicalDataset(
        PresentationChart chart,
        IReadOnlyList<SpreadsheetChartSeriesArtifact> series)
    {
        if (series.Count == 0) return false;
        if (chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble)
        {
            if (chart.Categories.Count != 0) return false;
            return series.All(item =>
                item.Values.Count > 0 &&
                item.Values.Count == item.XValues.Count &&
                (chart.Type != SpreadsheetChartType.Bubble || item.Values.Count == item.BubbleSizes.Count) &&
                (chart.Type != SpreadsheetChartType.Scatter || item.BubbleSizes.Count == 0) &&
                string.IsNullOrWhiteSpace(item.CategoryFormula) &&
                string.IsNullOrWhiteSpace(item.XValueFormula) &&
                string.IsNullOrWhiteSpace(item.ValueFormula) &&
                string.IsNullOrWhiteSpace(item.BubbleSizeFormula) &&
                item.MissingValueIndexes.Count == 0 &&
                item.Trendlines.Count == 0 && item.ErrorBars is null &&
                item.Fill is null && item.Line is null && item.Marker is null &&
                item.SeriesFill is null && item.DataLabels is null && !item.HasExplosion && !item.HasValuesFormatCode && item.PointStyles.Count == 0);
        }
        if (chart.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area or
            SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut or SpreadsheetChartType.Radar) ||
            chart.Categories.Count == 0)
            return false;
        return series.All(item =>
            item.Values.Count == chart.Categories.Count &&
            string.IsNullOrWhiteSpace(item.CategoryFormula) &&
            string.IsNullOrWhiteSpace(item.XValueFormula) &&
            string.IsNullOrWhiteSpace(item.ValueFormula) &&
            string.IsNullOrWhiteSpace(item.BubbleSizeFormula) &&
            item.XValues.Count == 0 && item.BubbleSizes.Count == 0 &&
            item.MissingValueIndexes.Count == 0 &&
            item.Trendlines.Count == 0 && item.ErrorBars is null &&
            item.Fill is null && item.Line is null && item.Marker is null &&
            item.SeriesFill is null && item.DataLabels is null && !item.HasExplosion && !item.HasValuesFormatCode && item.PointStyles.Count == 0);
    }

    private static JsonObject CanonicalDataset(
        PresentationChart chart,
        IReadOnlyList<SpreadsheetChartSeriesArtifact> series)
    {
        var rows = new JsonArray();
        if (chart.Type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble)
        {
            foreach (var item in series)
            {
                for (var pointIndex = 0; pointIndex < item.Values.Count; pointIndex++)
                {
                    var row = new JsonArray
                    {
                        NumberNode(item.XValues[pointIndex]),
                        StringNode(item.Name ?? string.Empty),
                        NumberNode(item.Values[pointIndex]),
                    };
                    if (chart.Type == SpreadsheetChartType.Bubble)
                        row.Add(NumberNode(item.BubbleSizes[pointIndex]));
                    rows.Add(row);
                }
            }
            return new JsonObject
            {
                ["cols"] = chart.Type == SpreadsheetChartType.Bubble
                    ? new JsonArray { StringNode("x"), StringNode("series"), StringNode("y"), StringNode("size") }
                    : new JsonArray { StringNode("x"), StringNode("series"), StringNode("y") },
                ["rows"] = rows,
            };
        }
        for (var categoryIndex = 0; categoryIndex < chart.Categories.Count; categoryIndex++)
        {
            foreach (var item in series)
            {
                var row = new JsonObject
                {
                    ["category"] = StringNode(chart.Categories[categoryIndex]),
                    ["series"] = StringNode(item.Name ?? string.Empty),
                };
                var missing = item.MissingValueIndexes.Contains((uint)categoryIndex);
                row["value"] = missing ? null : NumberNode(item.Values[categoryIndex]);
                rows.Add(row);
            }
        }
        return new JsonObject
        {
            ["cols"] = new JsonArray { StringNode("category"), StringNode("series"), StringNode("value") },
            ["rows"] = rows,
        };
    }

    private static void ProjectChartSeriesStyle(JsonObject output, SpreadsheetChartSeriesArtifact series)
    {
        if (series.SeriesFill is not null)
            output["fill"] = ProjectChartSurfaceFill(series.SeriesFill);
        else if (series.Fill is not null && !string.IsNullOrEmpty(series.Fill.Rgb))
            output["fill"] = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(series.Fill.Rgb)) };
        if (series.Line is not null && !string.IsNullOrEmpty(series.Line.Color?.Rgb))
            output["stroke"] = ProjectChartLine(series.Line);
        if (series.HasExplosion) output["explosion"] = JsonValue.Create(series.Explosion);
        if (series.HasValuesFormatCode) output["valuesFormatCode"] = StringNode(series.ValuesFormatCode);
        if (series.PointStyles.Count > 0)
        {
            var points = new JsonArray();
            foreach (var point in series.PointStyles.OrderBy(point => point.Index))
            {
            var item = new JsonObject { ["index"] = JsonValue.Create(point.Index) };
                if (point.Fill is not null) item["fill"] = ProjectChartSurfaceFill(point.Fill);
                if (point.Line is not null && !string.IsNullOrEmpty(point.Line.Color?.Rgb))
                    item["stroke"] = ProjectChartLine(point.Line);
                if (point.HasExplosion) item["explosion"] = JsonValue.Create(point.Explosion);
                points.Add(item);
            }
            output["pointStyles"] = points;
        }
        if (series.Marker is not null && series.Marker.Symbol != SpreadsheetChartMarkerSymbol.Unspecified)
        {
            var symbol = Marker(series.Marker.Symbol);
            if (symbol is not null)
            {
                if (!series.Marker.HasSize && series.Marker.Fill is null && series.Marker.Line is null)
                    output["marker"] = StringNode(symbol);
                else
                {
                    var marker = new JsonObject { ["symbol"] = StringNode(symbol) };
                    if (series.Marker.HasSize) marker["size"] = JsonValue.Create(series.Marker.Size);
                    if (series.Marker.Fill is not null && !string.IsNullOrEmpty(series.Marker.Fill.Rgb))
                    {
                        var color = Color(series.Marker.Fill.Rgb);
                        if (series.Marker.HasFillOpacityThousandthPercent)
                        {
                            var alpha = Math.Clamp((int)Math.Round(Unit(series.Marker.FillOpacityThousandthPercent) * 255), 0, 255);
                            color += $"{alpha:X2}";
                        }
                        marker["fill"] = StringNode(color);
                    }
                    if (series.Marker.Line is not null && !string.IsNullOrEmpty(series.Marker.Line.Color?.Rgb))
                        marker["stroke"] = ProjectChartLine(series.Marker.Line);
                    output["marker"] = marker;
                }
            }
        }
        if (series.Trendlines.Count > 0)
        {
            var trendlines = new JsonArray();
            foreach (var item in series.Trendlines)
            {
                if (TrendlineType(item.Type) is not { } type) continue;
                var trendline = new JsonObject { ["type"] = StringNode(type) };
                if (!string.IsNullOrEmpty(item.Name)) trendline["name"] = StringNode(item.Name);
                if (item.HasPolynomialOrder) trendline["order"] = JsonValue.Create(item.PolynomialOrder);
                if (item.HasPeriod) trendline["period"] = JsonValue.Create(item.Period);
                if (item.HasForward) trendline["forward"] = JsonValue.Create(item.Forward);
                if (item.HasBackward) trendline["backward"] = JsonValue.Create(item.Backward);
                if (item.HasIntercept) trendline["intercept"] = JsonValue.Create(item.Intercept);
                if (item.DisplayEquation) trendline["displayEquation"] = JsonValue.Create(true);
                if (item.DisplayRSquared) trendline["displayRSquared"] = JsonValue.Create(true);
                if (item.Line is not null && !string.IsNullOrEmpty(item.Line.Color?.Rgb))
                    trendline["stroke"] = ProjectChartLine(item.Line);
                if (item.Label is { } label)
                {
                    var value = new JsonObject();
                    if (label.Layout is not null) value["layout"] = OpenXmlChartLayoutCodec.Project(label.Layout);
                    if (label.HasText) value["text"] = StringNode(label.Text);
                    if (label.RichText is not null) value["text"] = OpenXmlChartRichTextCodec.Project(label.RichText, ProjectChartTextStyle);
                    if (label.HasNumberFormatCode) value["numberFormat"] = StringNode(label.NumberFormatCode);
                    if (label.NumberFormatLink == SpreadsheetChartNumberFormatLink.Source) value["numberFormatSourceLinked"] = JsonValue.Create(true);
                    if (label.NumberFormatLink == SpreadsheetChartNumberFormatLink.Omitted) value["numberFormatSourceLinked"] = null;
                    if (label.TextStyle is not null) value["textStyle"] = ProjectChartTextStyle(label.TextStyle);
                    if (label.Fill is not null) value["fill"] = ProjectChartSurfaceFill(label.Fill);
                    if (label.Line is not null) value["line"] = ProjectChartLine(label.Line);
                    trendline["label"] = value;
                }
                trendlines.Add(trendline);
            }
            if (trendlines.Count > 0) output["trendlines"] = trendlines;
        }
        if (series.ErrorBars is { } errorBars &&
            CanProjectErrorBarData(errorBars.Plus) && CanProjectErrorBarData(errorBars.Minus) &&
            ErrorBarDirection(errorBars.Direction) is { } direction &&
            ErrorBarType(errorBars.Type) is { } barType &&
            ErrorBarValueType(errorBars.ValueType) is { } valueType)
        {
            var projected = new JsonObject
            {
                ["direction"] = StringNode(direction),
                ["type"] = StringNode(barType),
                ["valueType"] = StringNode(valueType),
            };
            if (errorBars.HasValue) projected["value"] = JsonValue.Create(errorBars.Value);
            if (errorBars.Plus is not null) projected["plus"] = ProjectErrorBarData(errorBars.Plus);
            if (errorBars.Minus is not null) projected["minus"] = ProjectErrorBarData(errorBars.Minus);
            if (errorBars.NoEndCap) projected["noEndCap"] = JsonValue.Create(true);
            if (errorBars.Line is not null && !string.IsNullOrEmpty(errorBars.Line.Color?.Rgb))
                projected["stroke"] = ProjectChartLine(errorBars.Line);
            output["errorBars"] = projected;
        }
        if (series.DataLabels is not null)
            output["dataLabels"] = ProjectSeriesDataLabels(series.DataLabels);
    }

    private static JsonObject ProjectErrorBarData(SpreadsheetChartErrorBarDataArtifact source)
    {
        var data = new JsonObject { ["values"] = new JsonArray(source.Values.Select(value => (JsonNode?)JsonValue.Create(value)).ToArray()) };
        if (source.Formula.Length > 0) data["formula"] = StringNode(source.Formula);
        if (source.FormatCode.Length > 0) data["formatCode"] = StringNode(source.FormatCode);
        return data;
    }

    private static bool CanProjectErrorBarData(SpreadsheetChartErrorBarDataArtifact? source) =>
        source is null || source.Values.Count > 0 && (source.Formula.Length == 0 ||
            PptxChartErrorDataWorkbookCodec.TryParseRange(source.Formula, out var range) && range.Length == source.Values.Count);

    private static JsonObject ProjectSeriesDataLabels(SpreadsheetChartSeriesDataLabelsArtifact source)
    {
        var output = source.Defaults is null ? new JsonObject() : ProjectChartLabelOverride(source.Defaults);
        if (source.Points.Count > 0)
        {
            var points = new JsonArray();
            foreach (var point in source.Points.OrderBy(point => point.Index))
            {
                var item = ProjectChartLabelOverride(point.Override!);
                item["index"] = JsonValue.Create(point.Index);
                points.Add(item);
            }
            output["points"] = points;
        }
        return output;
    }

    private static JsonObject ProjectChartLabelOverride(SpreadsheetChartDataLabelOverrideArtifact source)
    {
        var output = new JsonObject();
        if (source.Text.Length > 0) output["text"] = StringNode(source.Text);
        if (source.HasShowValue) output["showValue"] = JsonValue.Create(source.ShowValue);
        if (source.HasShowCategoryName) output["showCategory"] = JsonValue.Create(source.ShowCategoryName);
        if (source.HasShowSeriesName) output["showSeries"] = JsonValue.Create(source.ShowSeriesName);
        if (source.HasShowPercent) output["showPercent"] = JsonValue.Create(source.ShowPercent);
        if (source.HasShowBubbleSize) output["showBubbleSize"] = JsonValue.Create(source.ShowBubbleSize);
        if (source.HasShowLeaderLines) output["showLeaderLines"] = JsonValue.Create(source.ShowLeaderLines);
        if (source.HasPosition && DataLabelPosition(source.Position) is { } position) output["position"] = StringNode(position);
        if (source.TextStyle is not null) output["textStyle"] = ProjectChartTextStyle(source.TextStyle);
        if (source.NumberFormatCode.Length > 0) output["numberFormat"] = StringNode(source.NumberFormatCode);
        if (source.Fill is not null) output["fill"] = ProjectChartSurfaceFill(source.Fill);
        if (source.Line is not null) output["line"] = ProjectChartLine(source.Line);
        return output;
    }

    private static JsonObject ProjectChartAxis(SpreadsheetChartAxisArtifact axis)
    {
        var output = new JsonObject();
        if (!string.IsNullOrEmpty(axis.Title)) output["title"] = StringNode(axis.Title);
        if (!string.IsNullOrEmpty(axis.NumberFormatCode)) output["numberFormat"] = StringNode(axis.NumberFormatCode);
        if (axis.HasTickLabelInterval) output["tickLabelInterval"] = JsonValue.Create(axis.TickLabelInterval);
        if (axis.HasMinimum) output["min"] = JsonValue.Create(axis.Minimum);
        if (axis.HasMaximum) output["max"] = JsonValue.Create(axis.Maximum);
        if (axis.HasMajorUnit) output["majorUnit"] = JsonValue.Create(axis.MajorUnit);
        if (axis.HasMinorUnit) output["minorUnit"] = JsonValue.Create(axis.MinorUnit);
        if (axis.HasLogBase) output["logBase"] = JsonValue.Create(axis.LogBase);
        if (axis.HasPosition) output["position"] = StringNode(axis.Position);
        if (axis.HasVisible) output["visible"] = JsonValue.Create(axis.Visible);
        if (axis.HasReverse) output["reverse"] = JsonValue.Create(axis.Reverse);
        if (axis.HasTickLabelsVisible) output["tickLabelsVisible"] = JsonValue.Create(axis.TickLabelsVisible);
        if (axis.HasTickLabelPosition) output["tickLabelPosition"] = StringNode(axis.TickLabelPosition);
        if (axis.HasMajorTickMark) output["majorTickMark"] = StringNode(axis.MajorTickMark);
        if (axis.HasMinorTickMark) output["minorTickMark"] = StringNode(axis.MinorTickMark);
        if (axis.AxisLine is not null && !string.IsNullOrEmpty(axis.AxisLine.Color?.Rgb))
            output["axisLine"] = ProjectChartLine(axis.AxisLine);
        else if (axis.HasAxisLineVisible)
            output["axisLine"] = JsonValue.Create(axis.AxisLineVisible);
        if (axis.AxisLine is { } axisLine &&
            (axisLine.StartArrow.Length > 0 || axisLine.EndArrow.Length > 0))
        {
            var arrows = new JsonObject();
            if (axisLine.StartArrow.Length > 0) arrows["start"] = StringNode(axisLine.StartArrow);
            if (axisLine.EndArrow.Length > 0) arrows["end"] = StringNode(axisLine.EndArrow);
            output["axisLineArrow"] = arrows;
        }
        if (axis.MajorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MajorGridlineStyle.Color?.Rgb))
            output["gridLine"] = ProjectChartLine(axis.MajorGridlineStyle);
        else if (axis.HasMajorGridlineVisible)
            output["gridLine"] = JsonValue.Create(axis.MajorGridlineVisible);
        else if (axis.HasShowMajorGridlines)
            output["gridLine"] = JsonValue.Create(axis.ShowMajorGridlines);
        if (axis.MinorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MinorGridlineStyle.Color?.Rgb))
            output["minorGridLine"] = ProjectChartLine(axis.MinorGridlineStyle);
        else if (axis.HasMinorGridlineVisible)
            output["minorGridLine"] = JsonValue.Create(axis.MinorGridlineVisible);
        else if (axis.HasShowMinorGridlines)
            output["minorGridLine"] = JsonValue.Create(axis.ShowMinorGridlines);
        if (axis.TextStyle is not null)
            output["textStyle"] = ProjectChartTextStyle(axis.TextStyle);
        if (axis.TitleTextStyle is not null)
            output["titleTextStyle"] = ProjectChartTextStyle(axis.TitleTextStyle);
        return output;
    }

    private static bool TryProjectRadarSpokeAxis(PresentationChart chart, out JsonObject output)
    {
        output = new JsonObject();
        var xAxis = chart.XAxis;
        var yAxis = chart.YAxis;
        if (xAxis is null || yAxis is null ||
            xAxis.Title.Length > 0 || xAxis.NumberFormatCode.Length > 0 || xAxis.HasTickLabelInterval ||
            xAxis.HasMinimum || xAxis.HasMaximum || xAxis.HasMajorUnit || xAxis.HasMinorUnit || xAxis.HasLogBase || xAxis.HasReverse && xAxis.Reverse ||
            xAxis.AxisLine is not null || xAxis.HasAxisLineVisible || xAxis.TextStyle is not null || xAxis.TitleTextStyle is not null ||
            xAxis.HasTickLabelsVisible || xAxis.HasTickLabelPosition ||
            xAxis.HasShowMinorGridlines || xAxis.HasMinorGridlineVisible || xAxis.MinorGridlineStyle is not null ||
            yAxis.Title.Length > 0 || yAxis.HasTickLabelInterval || yAxis.HasReverse && yAxis.Reverse ||
            yAxis.AxisLine is not null || yAxis.HasAxisLineVisible || yAxis.TitleTextStyle is not null ||
            yAxis.HasTickLabelPosition ||
            yAxis.HasShowMinorGridlines || yAxis.HasMinorGridlineVisible || yAxis.MinorGridlineStyle is not null)
            return false;

        if (xAxis.HasVisible != yAxis.HasVisible ||
            xAxis.HasVisible && xAxis.Visible != yAxis.Visible)
            return false;

        var hasEvidence = xAxis.HasVisible || yAxis.HasVisible ||
            xAxis.HasShowMajorGridlines || xAxis.HasMajorGridlineVisible || xAxis.MajorGridlineStyle is not null ||
            yAxis.HasShowMajorGridlines || yAxis.HasMajorGridlineVisible || yAxis.MajorGridlineStyle is not null ||
            yAxis.HasMinimum || yAxis.HasMaximum || yAxis.HasMajorUnit || yAxis.HasMinorUnit || yAxis.HasLogBase ||
            yAxis.HasTickLabelsVisible || yAxis.NumberFormatCode.Length > 0 || yAxis.TextStyle is not null;
        if (!hasEvidence) return false;

        var show = !xAxis.HasVisible || xAxis.Visible;
        if (xAxis.HasVisible) output["show"] = JsonValue.Create(show);
        if (yAxis.HasMinimum) output["min"] = JsonValue.Create(yAxis.Minimum);
        if (yAxis.HasMaximum) output["max"] = JsonValue.Create(yAxis.Maximum);
        if (yAxis.HasMajorUnit) output["majorUnit"] = JsonValue.Create(yAxis.MajorUnit);
        if (yAxis.HasMinorUnit) output["minorUnit"] = JsonValue.Create(yAxis.MinorUnit);
        if (yAxis.HasLogBase) output["logBase"] = JsonValue.Create(yAxis.LogBase);

        if (yAxis.HasTickLabelsVisible && !yAxis.TickLabelsVisible)
            output["label"] = JsonValue.Create(false);
        else if (yAxis.NumberFormatCode.Length > 0 || yAxis.TextStyle is not null)
        {
            var label = yAxis.TextStyle is null ? new JsonObject() : ProjectChartTextStyle(yAxis.TextStyle);
            if (yAxis.NumberFormatCode.Length > 0) label["numberFormat"] = StringNode(yAxis.NumberFormatCode);
            output["label"] = label;
        }
        else if (yAxis.HasTickLabelsVisible)
            output["label"] = JsonValue.Create(yAxis.TickLabelsVisible);

        if (show)
        {
            output["axisLine"] = ProjectRadarGuideLine(xAxis) ?? JsonValue.Create(false);
            output["gridLine"] = ProjectRadarGuideLine(yAxis) ?? JsonValue.Create(false);
        }
        return true;
    }

    private static JsonNode? ProjectRadarGuideLine(SpreadsheetChartAxisArtifact axis)
    {
        if (axis.MajorGridlineStyle is not null && !string.IsNullOrEmpty(axis.MajorGridlineStyle.Color?.Rgb))
            return ProjectChartLine(axis.MajorGridlineStyle);
        if (axis.HasMajorGridlineVisible) return JsonValue.Create(axis.MajorGridlineVisible);
        if (axis.HasShowMajorGridlines) return JsonValue.Create(axis.ShowMajorGridlines);
        return null;
    }

    private static JsonObject ProjectChartTextStyle(SpreadsheetChartTextStyleArtifact source)
    {
        var output = new JsonObject();
        if (source.HasFontSizePoints) output["fontSize"] = JsonValue.Create(source.FontSizePoints);
        if (source.FontFamily.Length > 0) output["fontFamily"] = StringNode(source.FontFamily);
        if (source.FontFamilyEastAsia.Length > 0) output["fontFamilyEastAsia"] = StringNode(source.FontFamilyEastAsia);
        if (source.FontFamilyComplexScript.Length > 0) output["fontFamilyComplexScript"] = StringNode(source.FontFamilyComplexScript);
        if (source.HasLanguage) output["language"] = StringNode(source.Language);
        if (source.HasStrike) output["strike"] = StringNode(source.Strike);
        if (source.HasBaselineThousandthPercent) output["baseline"] = JsonValue.Create(source.BaselineThousandthPercent / 1000d);
        if (source.HasLetterSpacingHundredthPoints) output["letterSpacing"] = JsonValue.Create(source.LetterSpacingHundredthPoints / 100d);
        if (source.HasKerningHundredthPoints) output["kerning"] = JsonValue.Create(source.KerningHundredthPoints / 100d);
        if (source.HasCapitalization) output["capitalization"] = StringNode(source.Capitalization);
        if (source.SoftEdge is not null) output["softEdge"] = SoftEdge(source.SoftEdge);
        if (source.Glow is not null) output["glow"] = Glow(source.Glow);
        if (source.Reflection is not null) output["reflection"] = ChartTextReflection(source.Reflection);
        if (source.InnerShadow is not null) output["innerShadow"] = ChartTextInnerShadow(source.InnerShadow);
        if (source.Shadow is not null) output["shadow"] = ChartTextShadow(source.Shadow);
        if (source.HasHighlightRgb) output["highlight"] = StringNode(Color(source.HighlightRgb));
        if (source.HasBold) output["bold"] = JsonValue.Create(source.Bold);
        if (source.HasItalic) output["italic"] = JsonValue.Create(source.Italic);
        if (source.Underline.Length > 0)
            output["underline"] = StringNode(source.Underline switch { "sng" => "single", "dbl" => "double", _ => source.Underline });
        if (source.Alignment.Length > 0)
            output["alignment"] = StringNode(source.Alignment switch { "l" => "left", "ctr" => "center", "r" => "right", "just" => "justify", _ => source.Alignment });
        if (source.Fill is not null)
            output["fill"] = ProjectChartSurfaceFill(source.Fill);
        if (source.ColorRgb.Length > 0)
            output["color"] = TextColor(source.ColorRgb, null,
                source.HasOpacityThousandthPercent, source.OpacityThousandthPercent);
        return output;
    }

    private static JsonObject ProjectChartLine(SpreadsheetChartLineStyleArtifact line)
    {
        var output = new JsonObject
        {
            ["color"] = StringNode(Color(line.Color.Rgb)),
            ["width"] = JsonValue.Create(line.HasWidthPoints ? line.WidthPoints : 0.75),
        };
        if (ChartDash(line.DashStyle) is { } dash) output["dash"] = StringNode(dash);
        if (line.Cap is "flat" or "round" or "square") output["cap"] = StringNode(line.Cap);
        if (line.Join is "miter" or "round" or "bevel") output["join"] = StringNode(line.Join);
        if (line.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(line.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectChartDataTable(SpreadsheetChartDataTableArtifact source)
    {
        var output = new JsonObject();
        if (source.HasShowHorizontalBorder) output["showHorizontalBorder"] = JsonValue.Create(source.ShowHorizontalBorder);
        if (source.HasShowVerticalBorder) output["showVerticalBorder"] = JsonValue.Create(source.ShowVerticalBorder);
        if (source.HasShowOutlineBorder) output["showOutlineBorder"] = JsonValue.Create(source.ShowOutlineBorder);
        if (source.HasShowLegendKey) output["showLegendKey"] = JsonValue.Create(source.ShowLegendKey);
        if (source.Fill is not null) output["fill"] = ProjectChartSurfaceFill(source.Fill);
        if (source.Line is not null && !string.IsNullOrEmpty(source.Line.Color?.Rgb))
            output["stroke"] = ProjectChartLine(source.Line);
        return output;
    }

    private static JsonObject ProjectChartSurfaceFill(SpreadsheetChartSurfaceFill fill)
    {
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            return new JsonObject { ["type"] = StringNode("none") };
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.GradientFill)
            return Gradient(fill.GradientFill);
        var output = new JsonObject
        {
            ["type"] = StringNode("solid"),
            ["color"] = StringNode(Color(fill.SolidRgb)),
        };
        if (fill.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(fill.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectTable(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var table = element.Table;
        if (table.ColumnWidthsEmu.Count == 0 || table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != table.ColumnWidthsEmu.Count))
            return ProjectOpaque(element, id, nativeRef, "table", "Preserved source table outside the bounded rectangular grid profile.");

        var output = ElementBase(id, element.Name, TableFrame(table), Accessibility(table.Accessibility), nativeRef);
        output["type"] = StringNode("table");
        var columns = new JsonArray();
        for (var index = 0; index < table.ColumnWidthsEmu.Count; index++)
            columns.Add(new JsonObject { ["id"] = StringNode($"column-{index + 1}"), ["width"] = JsonValue.Create(Points(table.ColumnWidthsEmu[index])) });
        output["columns"] = columns;
        output["rows"] = ProjectTableRows(table, context);
        var style = new JsonObject();
        if (table.HasFirstRow) style["headerRows"] = JsonValue.Create(table.FirstRow ? 1 : 0);
        if (table.HasBandedRows) style["bandedRows"] = JsonValue.Create(table.BandedRows);
        if (table.HasBandedColumns) style["bandedColumns"] = JsonValue.Create(table.BandedColumns);
        if (table.HasFirstColumn) style["firstColumnEmphasis"] = JsonValue.Create(table.FirstColumn);
        if (table.HasLastColumn) style["lastColumnEmphasis"] = JsonValue.Create(table.LastColumn);
        if (table.HasLastRow) style["lastRow"] = JsonValue.Create(table.LastRow);
        if (style.Count > 0) output["style"] = style;
        return output;
    }

    private static JsonArray ProjectTableRows(PresentationTable table, ProjectionContext context)
    {
        var mergeByOrigin = table.MergeRanges.ToDictionary(
            item => (Row: (int)item.StartRow, Column: (int)item.StartColumn));
        var covered = new HashSet<(int Row, int Column)>();
        foreach (var merge in table.MergeRanges)
            for (var row = (int)merge.StartRow; row <= merge.EndRow; row++)
                for (var column = (int)merge.StartColumn; column <= merge.EndColumn; column++)
                    if (row != merge.StartRow || column != merge.StartColumn) covered.Add((row, column));

        var rows = new JsonArray();
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var cells = new JsonArray();
            for (var columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++)
            {
                if (covered.Contains((rowIndex, columnIndex))) continue;
                var nativeCell = table.Rows[rowIndex].Cells[columnIndex];
                var cell = new JsonObject
                {
                    ["id"] = StringNode($"cell-{rowIndex + 1}-{columnIndex + 1}"),
                    // Keep the legacy string form for uniform/opaque cells.
                    // A bounded mixed-run cell carries its fixed-topology
                    // text body so each direct run style remains editable.
                    ["text"] = TextContent(nativeCell.TextBody, nativeCell.Text, context, includeBodyStyle: true),
                };
                if (mergeByOrigin.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    cell["rowSpan"] = JsonValue.Create(checked((int)(merge.EndRow - merge.StartRow + 1)));
                    cell["columnSpan"] = JsonValue.Create(checked((int)(merge.EndColumn - merge.StartColumn + 1)));
                }
                if (nativeCell.Fill is { } fill && ProjectTableCellFill(fill, context) is { } projectedFill)
                    cell["fill"] = projectedFill;
                if (nativeCell.Borders is { } borders)
                    cell["borders"] = ProjectTableCellBorders(borders);
                if (nativeCell.TextStyle is { } textStyle && ProjectTextStyle(textStyle) is { Count: > 0 } projectedTextStyle)
                    cell["textStyle"] = new JsonObject { ["defaultText"] = projectedTextStyle };
                cells.Add(cell);
            }
            rows.Add(new JsonObject
            {
                ["id"] = StringNode($"row-{rowIndex + 1}"),
                ["height"] = JsonValue.Create(Points(table.Rows[rowIndex].HeightEmu)),
                ["cells"] = cells,
            });
        }
        return rows;
    }

    private static JsonObject? ProjectTableCellFill(PresentationTableCellFill fill, ProjectionContext context)
    {
        if (fill.KindCase == PresentationTableCellFill.KindOneofCase.SolidRgb)
        {
            var output = new JsonObject
            {
                ["type"] = StringNode("solid"),
                ["color"] = StringNode(Color(fill.SolidRgb)),
            };
            if (fill.HasOpacityThousandthPercent) output["opacity"] = JsonValue.Create(Unit(fill.OpacityThousandthPercent));
            return output;
        }
        return fill.KindCase switch
        {
            PresentationTableCellFill.KindOneofCase.NoFill => new JsonObject { ["type"] = StringNode("none") },
            PresentationTableCellFill.KindOneofCase.GradientFill => Gradient(fill.GradientFill),
            PresentationTableCellFill.KindOneofCase.ImagePaint => ProjectImagePaint(fill.ImagePaint, context),
            _ => null,
        };
    }

    private static JsonObject ProjectTableCellBorders(PresentationTableCellBorders borders)
    {
        var output = new JsonObject();
        if (borders.Left is { } left) output["left"] = ProjectChartLine(left);
        if (borders.Top is { } top) output["top"] = ProjectChartLine(top);
        if (borders.Right is { } right) output["right"] = ProjectChartLine(right);
        if (borders.Bottom is { } bottom) output["bottom"] = ProjectChartLine(bottom);
        return output;
    }

    private static bool ProjectableCustomSite(PresentationSlide slide, string targetId, uint index)
    {
        bool Find(IEnumerable<PresentationElement> elements)
        {
            foreach (var element in elements)
            {
                if (element.Id == targetId)
                    return element.Source?.Editable == true && element.Shape is { Placeholder: null } shape &&
                        !PpjLinePathCodec.IsLineLike(shape) && CanProjectCustomGeometry(shape, allowShapeGraph: true) &&
                        index < shape.CustomConnectionSites.Count;
                if (element.Group is { } group && Find(group.Children)) return true;
            }
            return false;
        }
        return targetId.Length != 0 && Find(slide.Elements);
    }

    private static JsonObject ProjectConnector(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context,
        PresentationSlide slide)
    {
        var connector = element.Connector;
        var output = ElementBase(id, element.Name, ConnectorFrame(connector), Accessibility(connector.Accessibility), nativeRef);
        output["type"] = StringNode("connector");
        output["connectorType"] = StringNode(connector.ConnectorType is "elbow" or "curved" ? connector.ConnectorType : "straight");
        output["from"] = ConnectorEndpoint(connector.StartTargetId, connector.StartXEmu, connector.StartYEmu, pageId, context, connector.StartFrameAnchor,
            ProjectableCustomSite(slide, connector.StartTargetId, connector.StartConnectionSiteIndex) ? connector.StartConnectionSiteIndex : null);
        output["to"] = ConnectorEndpoint(connector.EndTargetId, connector.EndXEmu, connector.EndYEmu, pageId, context, connector.EndFrameAnchor,
            ProjectableCustomSite(slide, connector.EndTargetId, connector.EndConnectionSiteIndex) ? connector.EndConnectionSiteIndex : null);
        // A connector has one native line-alpha owner. Authored
        // compositing.opacity is therefore projected as the effective stroke
        // opacity rather than as a second, unrecoverable field.
        output["stroke"] = Stroke(
            connector.LineRgb,
            connector.LineWidthEmu,
            connector.LineStyle,
            connector.LineCap,
            connector.LineJoin,
            connector.HasLineOpacityThousandthPercent ? Unit(connector.LineOpacityThousandthPercent) : null,
            connector.LineScheme);
        if (Arrow(connector.StartArrow) is { } startArrow) output["startArrow"] = startArrow;
        if (Arrow(connector.EndArrow) is { } endArrow) output["endArrow"] = endArrow;
        if (connector.StartArrowWidth.Length > 0) output["startArrowWidth"] = connector.StartArrowWidth;
        if (connector.StartArrowLength.Length > 0) output["startArrowLength"] = connector.StartArrowLength;
        if (connector.EndArrowWidth.Length > 0) output["endArrowWidth"] = connector.EndArrowWidth;
        if (connector.EndArrowLength.Length > 0) output["endArrowLength"] = connector.EndArrowLength;
        if (connector.HasBendAdjustment) output["bendAdjustment"] = connector.BendAdjustment;
        return output;
    }

    private static JsonObject ProjectGroup(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        PresentationSlide slide,
        string pageId,
        ProjectionContext context,
        IReadOnlyList<uint> shapeTreePath)
    {
        var group = element.Group;
        if (group.Children.Count == 0)
            return ProjectOpaque(element, id, nativeRef, "group", "Preserved empty or unsupported native group.");
        var output = ElementBase(id, element.Name, GroupFrame(group), Accessibility(group.Accessibility), nativeRef);
        output["type"] = StringNode("group");
        output["childFrame"] = Frame(group.ChildLeftEmu, group.ChildTopEmu, group.ChildWidthEmu, group.ChildHeightEmu);
        var children = new JsonArray();
        for (var index = 0; index < group.Children.Count; index++)
        {
            var child = group.Children[index];
            children.Add(ProjectElement(
                child,
                slide,
                pageId,
                context,
                shapeTreePath.Concat([child.Source?.ShapeTreeIndex ?? checked((uint)index)]).ToArray()));
        }
        output["elements"] = children;
        if (children.Count > 0)
        {
            // DrawingML has no separate group-level accessibility order
            // owner in the bounded profile.  The local child shape-tree order
            // is therefore projected explicitly so a subsequent PPJ edit can
            // request a complete, auditable permutation without flattening
            // the group.
            var readingOrder = new JsonArray();
            foreach (var child in children)
            {
                if (child is JsonObject childObject && childObject["id"] is JsonValue childId)
                    // JsonArray.Add<T> requests runtime serialization metadata;
                    // use the same AOT-safe primitive as page reading order.
                    readingOrder.Add(StringNode(childId.GetValue<string>()));
            }
            if (readingOrder.Count == children.Count)
                output["readingOrder"] = readingOrder;
        }
        return output;
    }

    private static JsonObject ProjectOpaque(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string? kind = null,
        string? summary = null)
    {
        var opaque = element.ContentCase == PresentationElement.ContentOneofCase.Opaque ? element.Opaque : null;
        var nativeKind = kind ?? opaque?.NativeKind;
        if (string.IsNullOrWhiteSpace(nativeKind)) nativeKind = element.ContentCase.ToString();
        var output = ElementBase(id, element.Name, ElementFrame(element), Accessibility(opaque?.Accessibility), nativeRef);
        output["type"] = StringNode("opaque");
        output["nativeKind"] = StringNode(nativeKind);
        output["summary"] = StringNode(summary ?? $"Preserved source-owned {nativeKind} object; only issued nativeRef capabilities are editable.");
        if (opaque is not null && PpjNativeTextProjection.TryRead(opaque.RawXml, out var nativeLeaves))
        {
            var visible = new JsonArray();
            foreach (var leaf in nativeLeaves) visible.Add(StringNode(leaf));
            output["visibleText"] = visible;
        }
        else if (!string.IsNullOrWhiteSpace(opaque?.Text))
        {
            var visible = new JsonArray();
            foreach (var line in opaque.Text.Split('\n').Where(line => line.Length > 0)) visible.Add(StringNode(line));
            if (visible.Count > 0) output["visibleText"] = visible;
        }
        return output;
    }

    private static JsonObject ProjectSourceSmartArt(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context)
    {
        var diagram = element.Opaque.DiagramText;
        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = StringNode("smartArt");
        output["mode"] = StringNode("source-bound");
        if (context.SmartArtNativeSections(element.Opaque.PreservedPartPaths) is { } nativeSections)
            output["nativeSections"] = nativeSections;
        if (!string.IsNullOrWhiteSpace(diagram.LayoutDefinitionId))
            output["layoutDefinitionId"] = StringNode(diagram.LayoutDefinitionId);
        var nodes = new JsonArray();
        var nodeIds = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < diagram.Nodes.Count; index++)
        {
            var source = diagram.Nodes[index];
            var nodeId = context.UniqueId($"{id}-node-{NormalizeId(source.ModelId, $"node-{index + 1}")}");
            nodeIds.Add(source.ModelId, nodeId);
            var values = source.RunTexts.Count > 0 ? source.RunTexts.ToArray() : [source.Text];
            var hash = Sha256(Encoding.UTF8.GetBytes(source.ModelId + "\0" + string.Join("\0", values)));
            var capabilities = new[] { new CapabilitySpec("setSmartArtText", ["smartArt.text"]) };
            var node = new JsonObject
            {
                ["id"] = StringNode(nodeId),
                ["text"] = SmartArtText(values),
                ["nativeRef"] = NativeRef(
                    context,
                    $"element:{pageId}:{id}:smartArtNode:{index}",
                    hash,
                    capabilities),
            };
            node["kind"] = StringNode(source.PointType switch
            {
                "asst" => "assistant",
                "doc" => "document",
                _ => "node",
            });
            nodes.Add(node);
        }
        output["nodes"] = nodes;
        if (diagram.Connections.Count > 0)
        {
            var connections = new JsonArray();
            for (var index = 0; index < diagram.Connections.Count; index++)
            {
                var source = diagram.Connections[index];
                if (!nodeIds.TryGetValue(source.FromModelId, out var fromId) ||
                    !nodeIds.TryGetValue(source.ToModelId, out var toId))
                    throw new CodecException(
                        "ppj.smartArt.invalid_source_graph",
                        "A proven source-bound SmartArt parent edge no longer resolves to projected content nodes.",
                        $"elements.{id}.connections[{index}]");
                var connection = new JsonObject
                {
                    ["id"] = StringNode(context.UniqueId($"{id}-connection-{NormalizeId(source.ModelId, $"connection-{index + 1}")}")),
                    ["from"] = StringNode(fromId),
                    ["to"] = StringNode(toId),
                    ["role"] = StringNode("parent"),
                };
                if (source.Order > 0) connection["order"] = JsonValue.Create(source.Order);
                connections.Add(connection);
            }
            output["connections"] = connections;
        }
        return output;
    }

    private static JsonObject ProjectNativeSmartArt(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var diagram = element.Diagram;
        var output = ElementBase(id, element.Name, DiagramFrame(diagram), Accessibility(diagram.Accessibility), nativeRef);
        output["type"] = StringNode("smartArt");
        output["mode"] = StringNode("source-bound");
        if (!string.IsNullOrWhiteSpace(diagram.Layout)) output["layout"] = StringNode(diagram.Layout);
        if (!string.IsNullOrWhiteSpace(diagram.DefinitionAssetId) &&
            context.TryMaterializeAsset(diagram.DefinitionAssetId, out var definitionAssetId))
        {
            output.Remove("layout");
            output["definitionAsset"] = StringNode(definitionAssetId);
        }
        var nodes = new JsonArray();
        foreach (var node in diagram.Nodes)
        {
            var projected = new JsonObject
            {
                ["id"] = StringNode(node.Id),
                ["text"] = TextContent(node.TextBody, PptxTextCodec.Flatten(node.TextBody), context),
            };
            if (!string.IsNullOrWhiteSpace(node.AssetId) && context.TryMaterializeAsset(node.AssetId, out var assetId))
                projected["asset"] = StringNode(assetId);
            var cachedShape = diagram.Drawing?.Children
                .SingleOrDefault(child => child.Name == node.Id && child.Shape is not null)?.Shape;
            if (cachedShape?.ImageFill is { } imagePaint &&
                context.TryMaterializeAsset(imagePaint.AssetId, out _))
                projected["image"] = ProjectSmartArtNodeImage(imagePaint);
            nodes.Add(projected);
        }
        output["nodes"] = nodes;
        if (diagram.Connections.Count > 0)
        {
            var connections = new JsonArray();
            foreach (var connection in diagram.Connections)
            {
                var projected = new JsonObject
                {
                    ["id"] = StringNode(connection.Id),
                    ["from"] = StringNode(connection.FromId),
                    ["to"] = StringNode(connection.ToId),
                    ["role"] = StringNode(connection.Role),
                };
                if (connection.Order > 0) projected["order"] = JsonValue.Create(connection.Order);
                connections.Add(projected);
            }
            output["connections"] = connections;
        }
        return output;
    }

    private static JsonObject ProjectSmartArtNodeImage(PresentationImagePaint paint)
    {
        var output = new JsonObject
        {
            ["fit"] = StringNode(paint.Mode == PresentationImagePaint.Types.Mode.Tile ? "tile" : "stretch"),
        };
        if (paint.Crop is { } crop)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(crop.BottomThousandthPercent)),
            };
        }
        if (paint.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(paint.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject ProjectSourceOle(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        ProjectionContext context)
    {
        var workbook = element.Opaque.OleWorkbook;
        var package = element.Opaque.OleOfficePackage;
        var partPath = workbook?.PartPath ?? package?.PartPath;
        var contentType = workbook?.ContentType ?? package?.ContentType;
        if (string.IsNullOrWhiteSpace(partPath) || string.IsNullOrWhiteSpace(contentType) ||
            !context.TryMaterializeSourcePart(partPath, contentType, out var payloadAssetId))
            return ProjectOpaque(element, id, nativeRef, "oleObject", "Preserved source OLE object whose payload could not be materialized safely.");

        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = StringNode("ole");
        output["payloadAsset"] = payloadAssetId;
        return output;
    }

    private static JsonNode SmartArtText(IReadOnlyList<string> values)
    {
        if (values.Count == 1) return StringNode(values[0]);
        var runs = new JsonArray();
        for (var index = 0; index < values.Count; index++)
            runs.Add(new JsonObject { ["id"] = StringNode($"run-{index + 1}"), ["text"] = StringNode(values[index]) });
        return new JsonObject
        {
            ["paragraphs"] = new JsonArray
            {
                new JsonObject { ["id"] = StringNode("paragraph-1"), ["runs"] = runs },
            },
        };
    }

    private static JsonObject ElementBase(
        string id,
        string? name,
        JsonObject frame,
        JsonObject? accessibility,
        JsonObject nativeRef)
    {
        var output = new JsonObject
        {
            ["id"] = StringNode(id),
            ["frame"] = frame,
            ["nativeRef"] = nativeRef,
        };
        if (!string.IsNullOrWhiteSpace(name)) output["name"] = StringNode(name);
        if (accessibility is not null) output["accessibility"] = accessibility;
        return output;
    }

    private static JsonObject? ProjectAction(PresentationRunHyperlink? source, ProjectionContext context)
    {
        if (source is null) return null;
        var output = new JsonObject();
        switch (source.TargetCase)
        {
            case PresentationRunHyperlink.TargetOneofCase.Uri:
                output["uri"] = StringNode(source.Uri);
                break;
            case PresentationRunHyperlink.TargetOneofCase.SlideId:
                if (!context.TryPageId(source.SlideId, out var pageId)) return null;
                output["slide"] = StringNode(pageId);
                break;
            case PresentationRunHyperlink.TargetOneofCase.CustomShowId:
                if (!context.TryCustomShowId(source.CustomShowId, out var customShowId)) return null;
                output["customShow"] = StringNode(customShowId);
                if (source.HasReturnToSlide) output["returnToSlide"] = JsonValue.Create(source.ReturnToSlide);
                break;
            case PresentationRunHyperlink.TargetOneofCase.Action:
                output["verb"] = StringNode(source.Action);
                break;
            default:
                return null;
        }
        if (source.HasTooltip) output["tooltip"] = StringNode(source.Tooltip);
        if (source.HasTargetFrame) output["targetFrame"] = StringNode(source.TargetFrame);
        if (source.HasHistory) output["history"] = JsonValue.Create(source.History);
        if (source.HasHighlightClick) output["highlightClick"] = JsonValue.Create(source.HighlightClick);
        return output;
    }

    private static JsonObject ShapeStyle(
        PresentationShape shape,
        ProjectionContext context,
        bool omitOwnerOpacity = false)
    {
        var style = new JsonObject();
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(shape.FillRgb)) };
            if (!omitOwnerOpacity && shape.HasFillOpacityThousandthPercent)
                fill["opacity"] = JsonValue.Create(Unit(shape.FillOpacityThousandthPercent));
            style["fill"] = fill;
        }
        else if (shape.GradientFill is not null)
        {
            style["fill"] = Gradient(shape.GradientFill, includeOpacity: !omitOwnerOpacity);
        }
        else if (shape.ImageFill is not null && ProjectImagePaint(shape.ImageFill, context, includeOpacity: !omitOwnerOpacity) is { } imageFill)
        {
            style["fill"] = imageFill;
        }
        else if (!string.IsNullOrWhiteSpace(shape.ImageFillAssetId) &&
                 context.TryMaterializeAsset(shape.ImageFillAssetId, out var sourceImageAssetId))
        {
            style["fill"] = new JsonObject
            {
                ["type"] = StringNode("image"),
                ["asset"] = StringNode(sourceImageAssetId),
                ["fit"] = StringNode("stretch"),
            };
        }
        if ((!string.IsNullOrEmpty(shape.LineRgb) || !string.IsNullOrEmpty(shape.LineScheme)) && shape.LineStyle != "none")
            style["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin,
                !omitOwnerOpacity && shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                shape.LineScheme);
        if (shape.Shadow is not null && (!string.IsNullOrEmpty(shape.Shadow.ColorRgb) || !string.IsNullOrEmpty(shape.Shadow.ColorScheme)))
            style["shadow"] = Shadow(shape.Shadow, includeOpacity: !omitOwnerOpacity);
        if (shape.Glow is not null && (!string.IsNullOrEmpty(shape.Glow.ColorRgb) || !string.IsNullOrEmpty(shape.Glow.ColorScheme)))
            style["glow"] = Glow(shape.Glow);
        if (shape.InnerShadow is not null && (!string.IsNullOrEmpty(shape.InnerShadow.ColorRgb) || !string.IsNullOrEmpty(shape.InnerShadow.ColorScheme)))
            style["innerShadow"] = InnerShadow(shape.InnerShadow);
        if (shape.Reflection is not null)
            style["reflection"] = Reflection(shape.Reflection);
        if (shape.SoftEdge is not null)
            style["softEdge"] = SoftEdge(shape.SoftEdge);
        return style;
    }

    private static JsonNode TextContent(
        PresentationTextBody? body,
        string? fallback,
        ProjectionContext context,
        bool includeBodyStyle = false)
    {
        if (body is null || body.Paragraphs.Count == 0 ||
            body.Paragraphs.Any(paragraph => paragraph.Runs.Count == 0 ||
                paragraph.Runs.Any(run => run.ContentCase is not (PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.LineBreak or PresentationTextRun.ContentOneofCase.Field))))
            return StringNode(fallback ?? string.Empty);

        var paragraphs = new JsonArray();
        for (var paragraphIndex = 0; paragraphIndex < body.Paragraphs.Count; paragraphIndex++)
        {
            var source = body.Paragraphs[paragraphIndex];
            var paragraph = new JsonObject
            {
                ["id"] = StringNode($"paragraph-{paragraphIndex + 1}"),
            };
            var paragraphStyle = new JsonObject();
            if (source.HasLevel) paragraphStyle["level"] = JsonValue.Create(checked((int)source.Level));
            if (source.HasFontAlignment) paragraphStyle["fontAlignment"] = StringNode(source.FontAlignment);
            if (source.HasHangingPunctuation) paragraphStyle["hangingPunctuation"] = JsonValue.Create(source.HangingPunctuation);
            if (source.HasEastAsianLineBreak) paragraphStyle["eastAsianLineBreak"] = JsonValue.Create(source.EastAsianLineBreak);
            // Keep the signed endpoints valid; six-decimal point rounding can
            // otherwise move int.MinValue outside the public coordinate range.
            if (source.HasDefaultTabSizeEmu) paragraphStyle["defaultTabSize"] = JsonValue.Create(source.DefaultTabSizeEmu / EmuPerPoint);
            if (source.HasLatinLineBreak) paragraphStyle["latinLineBreak"] = JsonValue.Create(source.LatinLineBreak);
            if (source.HasRightToLeft) paragraphStyle["direction"] = StringNode(source.RightToLeft ? "right-to-left" : "left-to-right");
            if (source.HasAlignment && ParagraphAlignment(source.Alignment) is { } alignment)
                paragraphStyle["alignment"] = StringNode(alignment);
            if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                paragraphStyle["indent"] = JsonValue.Create(Points(source.MarginLeftEmu));
            if (source.RightMarginCase == PresentationTextParagraph.RightMarginOneofCase.MarginRightEmu)
                paragraphStyle["rightIndent"] = JsonValue.Create(Points(source.MarginRightEmu));
            if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                paragraphStyle["hanging"] = JsonValue.Create(-Points(source.IndentEmu));
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                paragraphStyle["lineSpacing"] = JsonValue.Create(Math.Max(0.001, source.LineSpacingPoints));
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingMultiplier)
                paragraphStyle["lineSpacingMultiplier"] = JsonValue.Create(Math.Max(0.00001, source.LineSpacingMultiplier));
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                paragraphStyle["spaceBefore"] = JsonValue.Create(Math.Max(0, source.SpaceBeforePoints));
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforeMultiplier)
                paragraphStyle["spaceBeforeMultiplier"] = JsonValue.Create(Math.Max(0, source.SpaceBeforeMultiplier));
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                paragraphStyle["spaceAfter"] = JsonValue.Create(Math.Max(0, source.SpaceAfterPoints));
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterMultiplier)
                paragraphStyle["spaceAfterMultiplier"] = JsonValue.Create(Math.Max(0, source.SpaceAfterMultiplier));
            if (source.TabStops.Count > 0)
            {
                var tabStops = new JsonArray();
                foreach (var tab in source.TabStops)
                    tabStops.Add(new JsonObject
                    {
                        ["position"] = JsonValue.Create(Points(tab.PositionEmu)),
                        ["alignment"] = StringNode(tab.Alignment),
                    });
                paragraphStyle["tabStops"] = tabStops;
            }
            if (source.HasNoTabStops && source.NoTabStops)
                paragraphStyle["noTabStops"] = true;
            if (source.DefaultRunStyleCase == PresentationTextParagraph.DefaultRunStyleOneofCase.DefaultRunProperties &&
                ProjectTextStyle(source.DefaultRunProperties) is { Count: > 0 } defaultText)
                paragraphStyle["defaultText"] = defaultText;
            if (ProjectBullet(source, context) is { } bullet) paragraphStyle["bullet"] = bullet;
            if (paragraphStyle.Count > 0) paragraph["style"] = paragraphStyle;

            var runs = new JsonArray();
            for (var runIndex = 0; runIndex < source.Runs.Count; runIndex++)
            {
                var sourceRun = source.Runs[runIndex];
                var run = new JsonObject { ["id"] = StringNode($"run-{paragraphIndex + 1}-{runIndex + 1}") };
                if (sourceRun.ContentCase == PresentationTextRun.ContentOneofCase.Field)
                {
                    var field = new JsonObject
                    {
                        ["type"] = StringNode(sourceRun.Field.Type),
                        ["text"] = StringNode(sourceRun.Field.Text),
                    };
                    if (PptxTextCodec.IsAutomaticFieldType(sourceRun.Field.Type)) field["automatic"] = true;
                    if (!string.IsNullOrWhiteSpace(sourceRun.Field.Id)) field["id"] = StringNode(sourceRun.Field.Id);
                    run["field"] = field;
                }
                else if (sourceRun.ContentCase == PresentationTextRun.ContentOneofCase.LineBreak)
                    run["break"] = true;
                else run["text"] = StringNode(sourceRun.Text);
                var style = RunStyle(sourceRun);
                if (style.Count > 0) run["style"] = style;
                if (sourceRun.HyperlinkCase == PresentationTextRun.HyperlinkOneofCase.RunHyperlink &&
                    sourceRun.RunHyperlink.TargetCase == PresentationRunHyperlink.TargetOneofCase.Uri &&
                    !string.IsNullOrWhiteSpace(sourceRun.RunHyperlink.Uri))
                    run["hyperlink"] = new JsonObject { ["uri"] = StringNode(sourceRun.RunHyperlink.Uri) };
                runs.Add(run);
            }
            paragraph["runs"] = runs;
            paragraphs.Add(paragraph);
        }
        var output = new JsonObject { ["paragraphs"] = paragraphs };
        if (includeBodyStyle && TextBoxStyle(body) is { Count: > 0 } bodyStyle)
            output["style"] = bodyStyle;
        return output;
    }

    private static JsonObject RunStyle(PresentationTextRun run)
    {
        var style = new JsonObject();
        if (run.HasFontFamily) style["fontFamily"] = StringNode(run.FontFamily);
        if (run.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = StringNode(run.FontFamilyEastAsia);
        if (run.HasFontFamilyComplexScript) style["fontFamilyComplexScript"] = StringNode(run.FontFamilyComplexScript);
        if (run.HasFontSizePoints && run.FontSizePoints > 0) style["size"] = JsonValue.Create(run.FontSizePoints);
        if (run.HasBold) style["bold"] = JsonValue.Create(run.Bold);
        if (run.HasItalic) style["italic"] = JsonValue.Create(run.Italic);
        if (run.HasColorRgb && !string.IsNullOrEmpty(run.ColorRgb))
            style["color"] = TextColor(run.ColorRgb, null, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.HasColorScheme && !string.IsNullOrEmpty(run.ColorScheme))
            style["color"] = TextColor(null, run.ColorScheme, run.HasColorOpacityThousandthPercent, run.ColorOpacityThousandthPercent);
        else if (run.GradientFill is not null)
            style["gradient"] = TextGradient(run.GradientFill);
        if (run.Shadow is not null) style["shadow"] = TextRunShadow(run.Shadow);
        if (run.Glow is not null) style["glow"] = Glow(run.Glow);
        if (run.InnerShadow is not null) style["innerShadow"] = InnerShadow(run.InnerShadow);
        if (run.Reflection is not null) style["reflection"] = Reflection(run.Reflection, includePositions: true, includeFadeAngle: true, includeScaleX: true, includeScaleY: true, includeSkewX: true, includeSkewY: true, includeAlignment: true, includeRotateWithShape: true);
        if (run.SoftEdge is not null) style["softEdge"] = SoftEdge(run.SoftEdge);
        if (run.HighlightCase == PresentationTextRun.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(run.HighlightRgb))
            style["highlight"] = StringNode(Color(run.HighlightRgb));
        if (run.HasUnderline) style["underline"] = StringNode(run.Underline switch { "sng" => "single", "dbl" => "double", _ => run.Underline });
        if (run.HasStrike) style["strike"] = JsonValue.Create(run.Strike);
        if (run.HasFontKerningPoints) style["kerning"] = JsonValue.Create(run.FontKerningPoints);
        if (run.HasFontBaselinePercent) style["baseline"] = JsonValue.Create(run.FontBaselinePercent);
        if (run.HasFontSpacingPoints) style["letterSpacing"] = JsonValue.Create(run.FontSpacingPoints);
        if (run.HasFontCaps) style["capitalization"] = StringNode(run.FontCaps);
        if (run.HasLanguage) style["language"] = StringNode(run.Language);
        return style;
    }

    private static JsonObject ProjectTextStyle(PresentationTextStyle source)
    {
        var style = new JsonObject();
        if (source.HasFontFamily) style["fontFamily"] = StringNode(source.FontFamily);
        if (source.HasFontFamilyEastAsia) style["fontFamilyEastAsia"] = StringNode(source.FontFamilyEastAsia);
        if (source.HasFontFamilyComplexScript) style["fontFamilyComplexScript"] = StringNode(source.FontFamilyComplexScript);
        if (source.HasFontSizePoints && source.FontSizePoints > 0) style["size"] = JsonValue.Create(source.FontSizePoints);
        if (source.HasBold) style["bold"] = JsonValue.Create(source.Bold);
        if (source.HasItalic) style["italic"] = JsonValue.Create(source.Italic);
        if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorRgb && !string.IsNullOrEmpty(source.ColorRgb))
            style["color"] = TextColor(source.ColorRgb, null, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.ColorCase == PresentationTextStyle.ColorOneofCase.ColorScheme && !string.IsNullOrEmpty(source.ColorScheme))
            style["color"] = TextColor(null, source.ColorScheme, source.HasColorOpacityThousandthPercent, source.ColorOpacityThousandthPercent);
        else if (source.GradientFill is not null)
            style["gradient"] = TextGradient(source.GradientFill);
        if (source.Shadow is not null) style["shadow"] = ChartTextShadow(source.Shadow);
        if (source.Glow is not null) style["glow"] = Glow(source.Glow);
        if (source.InnerShadow is not null) style["innerShadow"] = ChartTextInnerShadow(source.InnerShadow);
        if (source.Reflection is not null) style["reflection"] = ChartTextReflection(source.Reflection);
        if (source.SoftEdge is not null) style["softEdge"] = SoftEdge(source.SoftEdge);
        if (source.HighlightCase == PresentationTextStyle.HighlightOneofCase.HighlightRgb && !string.IsNullOrEmpty(source.HighlightRgb))
            style["highlight"] = StringNode(Color(source.HighlightRgb));
        else if (source.HighlightCase == PresentationTextStyle.HighlightOneofCase.HighlightScheme)
            style["highlight"] = new JsonObject { ["token"] = StringNode(source.HighlightScheme) };
        if (source.HasUnderline) style["underline"] = StringNode(source.Underline switch { "sng" => "single", "dbl" => "double", _ => source.Underline });
        if (source.HasStrike) style["strike"] = JsonValue.Create(source.Strike);
        if (source.HasFontKerningPoints) style["kerning"] = JsonValue.Create(source.FontKerningPoints);
        if (source.HasFontBaselinePercent) style["baseline"] = JsonValue.Create(source.FontBaselinePercent);
        if (source.HasFontSpacingPoints) style["letterSpacing"] = JsonValue.Create(source.FontSpacingPoints);
        if (source.HasFontCaps) style["capitalization"] = StringNode(source.FontCaps);
        if (source.HasLanguage) style["language"] = StringNode(source.Language);
        return style;
    }

    private static JsonObject? ProjectBullet(PresentationTextParagraph paragraph, ProjectionContext context)
    {
        JsonObject? bullet = paragraph.BulletCase switch
        {
            PresentationTextParagraph.BulletOneofCase.NoBullet => new JsonObject { ["type"] = StringNode("none") },
            PresentationTextParagraph.BulletOneofCase.BulletCharacter => new JsonObject
            {
                ["type"] = StringNode("character"),
                ["character"] = StringNode(paragraph.BulletCharacter),
            },
            PresentationTextParagraph.BulletOneofCase.AutoNumber => new JsonObject
            {
                ["type"] = StringNode("number"),
                ["scheme"] = StringNode(paragraph.AutoNumber.Scheme),
            },
            PresentationTextParagraph.BulletOneofCase.PictureBullet => ProjectPictureBullet(paragraph.PictureBullet, context),
            _ => null,
        };
        if (bullet is null) return null;
        // Font, color, and size metadata on a noBullet paragraph are inert
        // native residue. Do not emit them into the closed `none` PPJ shape;
        // the source-bound leaves still preserve and edit those tokens.
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.NoBullet)
            return bullet;
        if (paragraph.BulletCase == PresentationTextParagraph.BulletOneofCase.AutoNumber && paragraph.AutoNumber.HasStartAt)
            bullet["startAt"] = JsonValue.Create(checked((int)paragraph.AutoNumber.StartAt));
        if (paragraph.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFamily)
            bullet["fontFamily"] = StringNode(paragraph.BulletFontFamily);
        else if (paragraph.BulletFontCase == PresentationTextParagraph.BulletFontOneofCase.BulletFontFollowText)
            bullet["fontFollowText"] = JsonValue.Create(true);
        if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorRgb)
            bullet["color"] = BulletColor(
                paragraph.BulletColorRgb,
                null,
                paragraph.HasBulletColorOpacityThousandthPercent,
                paragraph.BulletColorOpacityThousandthPercent);
        else if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorScheme)
            bullet["color"] = TextColor(
                null,
                paragraph.BulletColorScheme,
                paragraph.HasBulletColorOpacityThousandthPercent,
                paragraph.BulletColorOpacityThousandthPercent);
        else if (paragraph.BulletColorCase == PresentationTextParagraph.BulletColorOneofCase.BulletColorFollowText)
            bullet["colorFollowText"] = JsonValue.Create(true);
        if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePoints)
            bullet["size"] = JsonValue.Create(paragraph.BulletSizePoints);
        else if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizePercent)
            bullet["sizePercent"] = JsonValue.Create(paragraph.BulletSizePercent);
        else if (paragraph.BulletSizeCase == PresentationTextParagraph.BulletSizeOneofCase.BulletSizeFollowText)
            bullet["sizeFollowText"] = JsonValue.Create(true);
        return bullet;
    }

    private static JsonObject? ProjectPictureBullet(PresentationPictureBullet picture, ProjectionContext context)
    {
        var bullet = new JsonObject { ["type"] = StringNode("picture") };
        if (picture.SourceCase == PresentationPictureBullet.SourceOneofCase.AssetId)
        {
            if (!context.TryMaterializeAsset(picture.AssetId, out var assetId)) return null;
            bullet["asset"] = StringNode(assetId);
        }
        else if (picture.SourceCase == PresentationPictureBullet.SourceOneofCase.Uri)
            bullet["uri"] = StringNode(picture.Uri);
        else
            return null;
        return bullet;
    }

    private static JsonArray ImportedThemeColors()
    {
        var output = new JsonArray();
        foreach (var item in new[]
        {
            (Id: "source-neutral", Value: "#000000"),
            (Id: "bg1", Value: "#FFFFFF"),
            (Id: "tx1", Value: "#000000"),
            (Id: "bg2", Value: "#EEECE1"),
            (Id: "tx2", Value: "#1F497D"),
            (Id: "accent1", Value: "#4F81BD"),
            (Id: "accent2", Value: "#C0504D"),
            (Id: "accent3", Value: "#9BBB59"),
            (Id: "accent4", Value: "#8064A2"),
            (Id: "accent5", Value: "#4BACC6"),
            (Id: "accent6", Value: "#F79646"),
            (Id: "hlink", Value: "#0000FF"),
            (Id: "folHlink", Value: "#800080"),
            (Id: "dk1", Value: "#000000"),
            (Id: "lt1", Value: "#FFFFFF"),
            (Id: "dk2", Value: "#1F497D"),
            (Id: "lt2", Value: "#EEECE1"),
        })
            output.Add(new JsonObject
            {
                ["id"] = StringNode(item.Id),
                ["value"] = StringNode(item.Value),
                ["role"] = StringNode("Fallback only; imported native styling remains source-owned"),
            });
        return output;
    }

    private static JsonObject TextBoxStyle(PresentationTextBody? body)
    {
        var output = new JsonObject();
        var properties = body?.BodyProperties;
        if (properties is null) return output;
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor)
            output["verticalAlignment"] = StringNode(properties.VerticalAnchor == "center" ? "middle" : properties.VerticalAnchor);
        if (properties.HasAnchorCenter)
            output["anchorCenter"] = JsonValue.Create(properties.AnchorCenter);
        if (properties.HasForceAntiAlias)
            output["forceAntiAlias"] = JsonValue.Create(properties.ForceAntiAlias);
        if (properties.HasSpaceFirstLastParagraph)
            output["spaceFirstLastParagraph"] = JsonValue.Create(properties.SpaceFirstLastParagraph);
        if (properties.HasCompatibleLineSpacing)
            output["compatibleLineSpacing"] = JsonValue.Create(properties.CompatibleLineSpacing);
        if (properties.HasFromWordArt)
            output["fromWordArt"] = JsonValue.Create(properties.FromWordArt);
        if (properties.HasTextWarpPreset)
            output["textWarpPreset"] = StringNode(properties.TextWarpPreset);
        if (properties.TextWarpAdjustments.Count > 0)
        {
            output["textWarpAdjustments"] = new JsonArray(properties.TextWarpAdjustments.Select(adjustment =>
                new JsonObject
                {
                    ["name"] = StringNode(adjustment.Name),
                    ["value"] = JsonValue.Create(adjustment.Value),
                }).ToArray());
        }
        if (properties.HasFlatTextZ)
            output["flatTextZ"] = JsonValue.Create(properties.FlatTextZ);
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap)
            output["wrap"] = StringNode(properties.Wrap);
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode)
            output["autoFit"] = StringNode(properties.AutoFitMode switch { "shrinkText" => "shrink-text", "resizeShape" => "resize-shape", _ => "none" });
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode &&
            properties.AutoFitMode == "shrinkText" && properties.NormalAutoFit is { } normalAutoFit &&
            (normalAutoFit.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000 ||
             normalAutoFit.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000))
        {
            var normal = new JsonObject();
            if (normalAutoFit.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000)
                normal["fontScale"] = JsonValue.Create(normalAutoFit.FontScale1000 / 1_000d);
            if (normalAutoFit.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000)
                normal["lineSpacingReduction"] = JsonValue.Create(normalAutoFit.LineSpacingReduction1000 / 1_000d);
            output["normalAutoFit"] = normal;
        }
        var margins = new JsonObject();
        if (properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) margins["left"] = JsonValue.Create(Points(properties.LeftInsetEmu));
        if (properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) margins["top"] = JsonValue.Create(Points(properties.TopInsetEmu));
        if (properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) margins["right"] = JsonValue.Create(Points(properties.RightInsetEmu));
        if (properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) margins["bottom"] = JsonValue.Create(Points(properties.BottomInsetEmu));
        if (margins.Count > 0) output["margins"] = margins;
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns)
            output["columns"] = JsonValue.Create(checked((int)properties.Columns));
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu)
            output["columnGap"] = JsonValue.Create(Points(properties.ColumnSpacingEmu));
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns)
            output["columnDirection"] = StringNode(properties.RightToLeftColumns ? "right-to-left" : "left-to-right");
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode)
            output["verticalText"] = StringNode(properties.VerticalTextMode);
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000)
            output["rotation"] = JsonValue.Create(properties.RotationAngle60000 / 60_000d);
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode)
            output["verticalOverflow"] = StringNode(properties.VerticalOverflowMode);
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode)
            output["horizontalOverflow"] = StringNode(properties.HorizontalOverflowMode);
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright)
            output["upright"] = JsonValue.Create(properties.Upright);
        return output;
    }

    private static string? ParagraphAlignment(string value) => value switch
    {
        "l" or "left" => "left",
        "ctr" or "center" => "center",
        "r" or "right" => "right",
        "just" or "justify" => "justify",
        "justLow" or "justifyLow" => "justifyLow",
        "dist" or "distributed" => "distributed",
        "thaiDist" or "thaiDistributed" => "thaiDistributed",
        _ => null,
    };

    private static void ApplyTextContainerStyle(JsonObject output, PresentationShape shape, ProjectionContext context)
    {
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(shape.FillRgb)) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = JsonValue.Create(Unit(shape.FillOpacityThousandthPercent));
            output["fill"] = fill;
        }
        else if (shape.GradientFill is not null)
        {
            output["fill"] = Gradient(shape.GradientFill);
        }
        else if (shape.ImageFill is not null && ProjectImagePaint(shape.ImageFill, context) is { } imageFill)
        {
            output["fill"] = imageFill;
        }
        if ((!string.IsNullOrEmpty(shape.LineRgb) || !string.IsNullOrEmpty(shape.LineScheme)) && shape.LineStyle != "none")
            output["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin,
                shape.HasLineOpacityThousandthPercent ? Unit(shape.LineOpacityThousandthPercent) : null,
                shape.LineScheme);
    }

    private static JsonObject? ProjectBackground(PresentationBackground? background, ProjectionContext context)
    {
        if (background is null) return null;
        if (background.ImagePaint is not null)
            return ProjectImagePaint(background.ImagePaint, context);
        if (!string.IsNullOrEmpty(background.ImageAssetId) && context.TryMaterializeAsset(background.ImageAssetId, out var assetId))
            return new JsonObject { ["type"] = StringNode("image"), ["asset"] = StringNode(assetId), ["fit"] = StringNode("stretch") };
        if (background.GradientFill is not null)
            return Gradient(background.GradientFill);
        if (!string.IsNullOrEmpty(background.ColorRgb))
        {
            var output = new JsonObject { ["type"] = StringNode("solid"), ["color"] = StringNode(Color(background.ColorRgb)) };
            if (background.HasOpacityThousandthPercent)
                output["opacity"] = JsonValue.Create(Unit(background.OpacityThousandthPercent));
            return output;
        }
        return null;
    }

    private static JsonObject? ProjectImagePaint(
        PresentationImagePaint paint,
        ProjectionContext context,
        bool includeOpacity = true)
    {
        if (!context.TryMaterializeAsset(paint.AssetId, out var assetId)) return null;
        var output = new JsonObject
        {
            ["type"] = StringNode("image"),
            ["asset"] = StringNode(assetId),
            ["fit"] = StringNode(paint.Mode == PresentationImagePaint.Types.Mode.Tile ? "tile" : "stretch"),
        };
        if (paint.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = JsonValue.Create(Crop(paint.Crop.LeftThousandthPercent)),
                ["top"] = JsonValue.Create(Crop(paint.Crop.TopThousandthPercent)),
                ["right"] = JsonValue.Create(Crop(paint.Crop.RightThousandthPercent)),
                ["bottom"] = JsonValue.Create(Crop(paint.Crop.BottomThousandthPercent)),
            };
        }
        if (includeOpacity && paint.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(paint.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject Gradient(PresentationGradientFill source, bool includeOpacity = true)
    {
        var stops = new JsonArray();
        foreach (var stop in source.Stops)
        {
            var item = new JsonObject
            {
                ["offset"] = JsonValue.Create(Unit(stop.PositionThousandthPercent)),
                ["color"] = StringNode(Color(stop.ColorRgb)),
            };
            if (includeOpacity && stop.HasOpacityThousandthPercent)
                item["opacity"] = JsonValue.Create(Unit(stop.OpacityThousandthPercent));
            stops.Add(item);
        }
        var output = new JsonObject
        {
            ["type"] = StringNode("gradient"),
            ["kind"] = StringNode(source.Kind == PresentationGradientFill.Types.Kind.Radial ? "radial" : "linear"),
            ["stops"] = stops,
        };
        if (source.Kind == PresentationGradientFill.Types.Kind.Linear && source.HasAngle60000)
            output["angle"] = JsonValue.Create(source.Angle60000 / 60_000d);
        return output;
    }

    private static JsonObject TextGradient(PresentationGradientFill source)
    {
        var output = Gradient(source);
        output.Remove("type");
        return output;
    }

    private static JsonArray ProjectAnimations(PresentationSlide slide, ProjectionContext context)
    {
        var output = new JsonArray();
        var pageId = context.PageId(slide.Id);
        foreach (var animation in slide.Animations)
        {
            if (!context.TryElementId(pageId, animation.TargetId, out var targetId)) continue;
            var item = new JsonObject
            {
                ["id"] = StringNode(context.UniqueId($"animation-{animation.Id}")),
                ["target"] = StringNode(targetId),
                ["phase"] = StringNode(animation.Phase),
                ["effect"] = StringNode(animation.Effect),
                ["start"] = StringNode(animation.Start),
                // PresentationML exposes the effective start condition, not
                // the PPJ sugar distinction.  Emit the normalized trigger as
                // well so imported click/previous timing remains a typed,
                // inspectable field instead of disappearing during projection.
                ["trigger"] = StringNode(animation.Start),
                ["durationMs"] = JsonValue.Create(animation.HasDurationMs ? checked((int)animation.DurationMs) : 500),
            };
            if (!string.IsNullOrEmpty(animation.Direction)) item["direction"] = StringNode(animation.Direction);
            if (animation.HasDelayMs) item["delayMs"] = JsonValue.Create(checked((int)animation.DelayMs));
            if (!string.IsNullOrEmpty(animation.TextBuild)) item["textBuild"] = StringNode(animation.TextBuild);
            if (!string.IsNullOrEmpty(animation.ChartBuild)) item["chartBuild"] = StringNode(animation.ChartBuild);
            if (animation.HasStaggerMs) item["staggerMs"] = JsonValue.Create(checked((int)animation.StaggerMs));
            if (animation.HasAnimateChartBackground) item["animateChartBackground"] = JsonValue.Create(animation.AnimateChartBackground);
            if (animation.HasRepeatCount) item["repeat"] = JsonValue.Create(checked((int)animation.RepeatCount));
            if (animation.HasAutoReverse) item["autoReverse"] = JsonValue.Create(animation.AutoReverse);
            if (!string.IsNullOrEmpty(animation.Easing) && animation.Easing != "linear") item["easing"] = StringNode(animation.Easing);
            output.Add(item);
        }
        return output;
    }

    private static JsonObject? ProjectTransition(
        PresentationSlide slide,
        PresentationArtifact presentation,
        ProjectionContext context)
    {
        if (slide.Morph is not null && slide.Morph.Pairs.Count > 0)
        {
            var fromPage = presentation.Slides.FirstOrDefault(item => item.Id == slide.Morph.FromSlideId);
            if (fromPage is null) return null;
            var pairs = new JsonArray();
            foreach (var pair in slide.Morph.Pairs)
            {
                if (!context.TryElementId(context.PageId(fromPage.Id), pair.FromId, out var fromId) ||
                    !context.TryElementId(context.PageId(slide.Id), pair.ToId, out var toId)) continue;
                pairs.Add(new JsonObject { ["key"] = StringNode(context.UniqueId($"morph-{pair.Key}")), ["from"] = StringNode(fromId), ["to"] = StringNode(toId) });
            }
            if (pairs.Count == 0) return null;
            return new JsonObject
            {
                ["type"] = StringNode("morph"),
                ["durationMs"] = JsonValue.Create(slide.Morph.HasDurationMs ? checked((int)slide.Morph.DurationMs) : 800),
                ["fromPage"] = StringNode(context.PageId(fromPage.Id)),
                ["morphPairs"] = pairs,
            };
        }
        var transition = slide.Transition;
        if (transition is null || !PpjTransitionLowering.IsBaseEffect(transition.Effect)) return null;
        var output = new JsonObject
        {
            ["type"] = StringNode(transition.Effect),
            ["speed"] = StringNode(transition.Speed),
            ["advanceOnClick"] = JsonValue.Create(transition.AdvanceOnClick),
        };
        if (transition.HasDurationMs) output["durationMs"] = JsonValue.Create(checked((int)transition.DurationMs));
        if (!string.IsNullOrEmpty(transition.Direction)) output["direction"] = StringNode(transition.Direction);
        if (!string.IsNullOrEmpty(transition.Orientation)) output["orientation"] = StringNode(transition.Orientation);
        if (transition.HasThroughBlack) output["throughBlack"] = JsonValue.Create(transition.ThroughBlack);
        if (transition.HasSpokes) output["spokes"] = JsonValue.Create(checked((int)transition.Spokes));
        if (transition.HasAdvanceAfterMs) output["advanceAfterMs"] = JsonValue.Create(checked((int)transition.AdvanceAfterMs));
        return output;
    }

    private static JsonArray ProjectSections(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        for (var sectionIndex = 0; sectionIndex < presentation.Sections.Count; sectionIndex++)
        {
            var section = presentation.Sections[sectionIndex];
            var pages = new JsonArray();
            foreach (var id in section.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            var item = new JsonObject
            {
                ["id"] = StringNode(context.UniqueId($"section-{section.Id}")),
                ["name"] = StringNode(string.IsNullOrWhiteSpace(section.Name) ? "Section" : section.Name),
                ["pages"] = pages,
            };
            if (section.Source is { } source)
            {
                var capabilities = source.Editable
                    ? new[] { new CapabilitySpec("setName", ["name"]), new CapabilitySpec("setPages", ["pages"]) }
                    : [];
                item["nativeRef"] = NativeRef(
                    context,
                    $"section:{sectionIndex}",
                    HashOrFallback(source.SectionXmlSha256, section),
                    capabilities);
            }
            output.Add(item);
        }
        return output;
    }

    private static JsonArray ProjectCustomShows(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        for (var showIndex = 0; showIndex < presentation.CustomShows.Count; showIndex++)
        {
            var show = presentation.CustomShows[showIndex];
            var pages = new JsonArray();
            foreach (var id in show.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(StringNode(pageId));
            if (pages.Count == 0) continue;
            var item = new JsonObject
            {
                ["id"] = StringNode(context.CustomShowId(show.Id)),
                ["name"] = StringNode(string.IsNullOrWhiteSpace(show.Name) ? "Custom show" : show.Name),
                ["pages"] = pages,
            };
            if (show.Source is { } source)
            {
                var capabilities = source.Editable
                    ? new[] { new CapabilitySpec("setName", ["name"]), new CapabilitySpec("setPages", ["pages"]) }
                    : [];
                item["nativeRef"] = NativeRef(
                    context,
                    $"customShow:{showIndex}",
                    HashOrFallback(source.ShowXmlSha256, show),
                    capabilities);
            }
            output.Add(item);
        }
        return output;
    }

    private static JsonArray ProjectComments(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.PageId(slide.Id);
            for (var commentIndex = 0; commentIndex < slide.LegacyComments.Count; commentIndex++)
            {
                var comment = slide.LegacyComments[commentIndex];
                if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Text)) continue;
                var capabilities = slide.Source?.LegacyCommentsEditable == true
                    ? new[] { new CapabilitySpec("replaceText", ["text"]) }
                    : [];
                var commentHash = HashOrFallback(null, comment);
                output.Add(new JsonObject
                {
                    ["id"] = StringNode(context.UniqueId($"comment-{pageId}-{comment.Id}")),
                    ["page"] = StringNode(pageId),
                    ["author"] = StringNode(string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author),
                    ["text"] = StringNode(comment.Text),
                    ["createdAt"] = StringNode(createdAt.ToUniversalTime().ToString("O")),
                    ["resolved"] = JsonValue.Create(false),
                    ["position"] = new JsonObject { ["x"] = JsonValue.Create(Points(comment.PositionXEmu)), ["y"] = JsonValue.Create(Points(comment.PositionYEmu)) },
                    ["nativeRef"] = NativeRef(context, $"comment:{pageId}:{commentIndex}", commentHash, capabilities),
                });
            }

            for (var threadIndex = 0; threadIndex < slide.ModernComments.Count; threadIndex++)
            {
                var thread = slide.ModernComments[threadIndex];
                if (thread.Root is null || thread.Anchor is null || thread.Anchor.Monikers.Count != 1)
                    continue;

                var sourceTargetId = thread.TargetId;
                if (sourceTargetId.EndsWith("/text", StringComparison.Ordinal))
                    sourceTargetId = sourceTargetId[..^5];
                if (!context.TryElementId(pageId, sourceTargetId, out var targetId))
                    continue;

                var rootId = context.UniqueId($"comment-{pageId}-modern-{threadIndex}-{thread.Root.Id}");
                ProjectModernComment(
                    output,
                    context,
                    pageId,
                    targetId,
                    thread,
                    thread.Root,
                    rootId,
                    parentId: null,
                    threadIndex,
                    replyIndex: null,
                    includeAnchor: true,
                    includePosition: true);
                for (var replyIndex = 0; replyIndex < thread.Replies.Count; replyIndex++)
                {
                    var reply = thread.Replies[replyIndex];
                    ProjectModernComment(
                        output,
                        context,
                        pageId,
                        targetId,
                        thread,
                        reply,
                        context.UniqueId($"comment-{pageId}-modern-{threadIndex}-reply-{replyIndex}-{reply.Id}"),
                        rootId,
                        threadIndex,
                        replyIndex,
                        includeAnchor: false,
                        includePosition: false);
                }
            }
        }
        return output;
    }

    private static void ProjectModernComment(
        JsonArray output,
        ProjectionContext context,
        string pageId,
        string targetId,
        PresentationModernCommentThread thread,
        PresentationModernComment comment,
        string commentId,
        string? parentId,
        int threadIndex,
        int? replyIndex,
        bool includeAnchor,
        bool includePosition)
    {
        if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Status))
            return;

        var capabilities = thread.Source?.Editable == true
            ? new[]
            {
                new CapabilitySpec("replaceText", ["text"]),
                new CapabilitySpec("setCommentStatus", ["status", "resolved"]),
            }
            : [];
        var item = new JsonObject
        {
            ["id"] = StringNode(commentId),
            ["page"] = StringNode(pageId),
            ["kind"] = StringNode("modern"),
            ["author"] = StringNode(string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author),
            ["text"] = StringNode(comment.Text),
            ["createdAt"] = StringNode(createdAt.ToUniversalTime().ToString("O")),
            ["resolved"] = JsonValue.Create(!comment.Status.Equals("active", StringComparison.Ordinal)),
            ["status"] = StringNode(comment.Status),
            ["nativeRef"] = NativeRef(
                context,
                replyIndex is { } index
                    ? $"comment:{pageId}:modern:{threadIndex}:reply:{index}"
                    : $"comment:{pageId}:modern:{threadIndex}:root",
                HashOrFallback(null, comment),
                capabilities),
        };
        if (parentId is not null) item["parent"] = StringNode(parentId);
        else item["target"] = StringNode(targetId);
        if (includePosition)
        {
            item["position"] = new JsonObject
            {
                ["x"] = JsonValue.Create(Points(thread.PositionXEmu)),
                ["y"] = JsonValue.Create(Points(thread.PositionYEmu)),
            };
        }
        if (includeAnchor)
        {
            var anchor = new JsonObject
            {
                ["kind"] = StringNode(thread.Anchor.Kind == PresentationModernCommentAnchor.Types.Kind.TextRange ? "textRange" : "element"),
                ["moniker"] = StringNode(thread.Anchor.Monikers[0].Type),
            };
            if (thread.Anchor.HasTextStart) anchor["textStart"] = JsonValue.Create(thread.Anchor.TextStart);
            if (thread.Anchor.HasTextLength) anchor["textLength"] = JsonValue.Create(thread.Anchor.TextLength);
            if (thread.Anchor.HasContextLength) anchor["contextLength"] = JsonValue.Create(thread.Anchor.ContextLength);
            if (thread.Anchor.HasContextHash) anchor["contextHash"] = JsonValue.Create(thread.Anchor.ContextHash);
            item["anchor"] = anchor;
        }
        output.Add(item);
    }

    private static IReadOnlyList<CapabilitySpec> Capabilities(
        PresentationElement element,
        bool hasEffectivePlaceholderFrame = false,
        bool customSiteEndpoints = false)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        // Accessibility is a separate non-visual cNvPr leaf.  Keep its
        // capability independent from the visual writer so a source-bound
        // edit can change title/description/decorative state without
        // widening the element's paint, geometry, or text profile.
        if (source.AccessibilityEditable)
            output.Add(new("setAccessibility", ["accessibility"]));
        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                if (source.TextEditable && TextTopologyRepresentable(element.Shape.TextBody))
                {
                    output.Add(new("replaceText", ["text"]));
                    // Slide placeholders use the text-only source-bound
                    // writer; it intentionally rejects paragraph/body-style
                    // changes because those may be inherited from the
                    // layout/master. Keep those capabilities on ordinary
                    // shape/text owners only.
                    if (element.Shape.Placeholder is null)
                    {
                        output.Add(new("setTextParagraphStyle", [
                            "text.paragraphs[].style.alignment",
                            "text.paragraphs[].style.level",
                            "text.paragraphs[].style.defaultTabSize",
                            "text.paragraphs[].style.eastAsianLineBreak",
                            "text.paragraphs[].style.latinLineBreak",
                            "text.paragraphs[].style.hangingPunctuation",
                            "text.paragraphs[].style.fontAlignment",
                            "text.paragraphs[].style.direction",
                            "text.paragraphs[].style.bullet.startAt",
                            "text.paragraphs[].style.bullet.character",
                            "text.paragraphs[].style.bullet.fontFamily",
                            "text.paragraphs[].style.bullet.fontFollowText",
                            "text.paragraphs[].style.bullet.color",
                            "text.paragraphs[].style.bullet.colorFollowText",
                            "text.paragraphs[].style.bullet.size",
                            "text.paragraphs[].style.bullet.sizePercent",
                            "text.paragraphs[].style.bullet.sizeFollowText",
                            "text.paragraphs[].style.bullet.scheme",
                            "text.paragraphs[].style.bullet.format",
                            "text.paragraphs[].style.tabStops",
                            "text.paragraphs[].style.noTabStops",
                            "text.paragraphs[].style.spaceBefore",
                            "text.paragraphs[].style.spaceBeforeMultiplier",
                            "text.paragraphs[].style.spaceAfter",
                            "text.paragraphs[].style.spaceAfterMultiplier",
                            "text.paragraphs[].style.lineSpacing",
                            "text.paragraphs[].style.lineSpacingMultiplier",
                            "text.paragraphs[].style.indent",
                            "text.paragraphs[].style.rightIndent",
                            "text.paragraphs[].style.hanging",
                            "text.paragraphs[].style.defaultText.bold",
                            "text.paragraphs[].style.defaultText.italic",
                            "text.paragraphs[].style.defaultText.size",
                            "text.paragraphs[].style.defaultText.fontFamily",
                            "text.paragraphs[].style.defaultText.fontFamilyEastAsia",
                            "text.paragraphs[].style.defaultText.fontFamilyComplexScript",
                            "text.paragraphs[].style.defaultText.language",
                            "text.paragraphs[].style.defaultText.kerning",
                            "text.paragraphs[].style.defaultText.letterSpacing",
                            "text.paragraphs[].style.defaultText.baseline",
                            "text.paragraphs[].style.defaultText.capitalization",
                            "text.paragraphs[].style.defaultText.strike",
                            "text.paragraphs[].style.defaultText.underline",
                            "text.paragraphs[].style.defaultText.highlight",
                            "text.paragraphs[].style.defaultText.color",
                            "text.paragraphs[].style.defaultText.gradient",
                            "text.paragraphs[].style.defaultText.glow",
                            "text.paragraphs[].style.defaultText.innerShadow",
                            "text.paragraphs[].style.defaultText.reflection",
                            "text.paragraphs[].style.defaultText.shadow",
                            "text.paragraphs[].style.defaultText.softEdge",
                        ]));
                        var staticFields = element.Shape.TextBody?.Paragraphs
                            .SelectMany(paragraph => paragraph.Runs)
                            .Where(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Field &&
                                !PptxTextCodec.IsAutomaticFieldType(run.Field.Type))
                            .ToArray() ?? [];
                        if (staticFields.Length > 0)
                        {
                            var fields = new List<string> { "text.paragraphs[].runs[].field.type" };
                            if (staticFields.Any(run => PptxTextCodec.ValidFieldId(run.Field.Id)))
                                fields.Add("text.paragraphs[].runs[].field.id");
                            output.Add(new("setTextField", fields));
                        }
                        if (PptxBodyPropertiesCodec.SupportsBoundedDirectLayout(element.Shape.TextBody?.BodyProperties))
                        {
                            output.Add(new("setTextBodyStyle", [
                                element.Shape.Geometry is not ("textbox" or "none" or "")
                                    ? "textStyle"
                                    : "text.style",
                            ]));
                        }
                    }
                }
                if (element.Shape.Placeholder is not null &&
                    source.DirectFramePresenceEditable &&
                    (element.Shape.DirectFrame is not null || hasEffectivePlaceholderFrame))
                    // A complete direct transform is owner-local.  When the
                    // slide only inherits its geometry, the same fields are
                    // safe to materialize on the slide after the effective
                    // frame has been uniquely resolved from layout/master.
                    output.Add(new("setFrame", EditableFrameFields));
                if (source.Editable)
                {
                    // A recognized shape-tree click action is owned by the
                    // same cNvPr closure as the shape.  Advertise a narrow
                    // replacement capability; PptxHyperlinkCodec still
                    // rejects unknown sound/macro/extension children.
                    output.Add(new("setAction", ["action"]));
                    output.Add(new("setHoverAction", ["hoverAction"]));
                    if (element.Shape.Geometry == "line" || PpjLinePathCodec.IsLineLike(element.Shape))
                    {
                        output.Add(new("setLinePath", ["line.path"]));
                        output.Add(new("setStroke", ["stroke"]));
                        output.Add(new("setShapeEffects", ["shape.shadow", "shape.glow", "shape.innerShadow", "shape.reflection", "shape.softEdge"]));
                        output.Add(new("setFrame", EditableFrameFields));
                        break;
                    }
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or "" or "custom") &&
                        TryGetCompoundShapeOpacity(element.Shape, out _))
                        output.Add(new("setOpacity", ["compositing.opacity"]));
                    // Source-bound image-filled custom geometry is projected
                    // for discovery and frame/stroke edits, but its native
                    // fill graph is not represented by a lossless PPJ
                    // replacement operation. Do not advertise setFill for
                    // that bounded shape profile.
                    if (string.IsNullOrWhiteSpace(element.Shape.ImageFillAssetId) ||
                        element.Shape.Geometry != "custom")
                        output.Add(new("setFill", ["fill"]));
                    output.Add(new("setStroke", ["stroke"]));
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or ""))
                        output.Add(new("setShapeEffects", ["shape.shadow", "shape.glow", "shape.innerShadow", "shape.reflection", "shape.softEdge"]));
                    output.Add(new("setFrame", element.Shape.Placeholder is null ? EditableFrameFields : PositionFrameFields));
                    if (element.Shape.Placeholder is null &&
                        element.Shape.Geometry is not ("textbox" or "none" or "custom") &&
                        PptxPresetGeometryAdjustmentCodec.TryExpectedCount(element.Shape.Geometry, out var adjustmentCount) &&
                        adjustmentCount > 0)
                        output.Add(new("setGeometry", ["geometry.adjustments"]));
                    else if (element.Shape.Placeholder is null &&
                             element.Shape.Geometry == "custom" &&
                             CanProjectCustomGeometry(element.Shape, allowShapeGraph: true))
                        output.Add(new("setGeometry", ["geometry.paths", "geometry.textRectangle", "geometry.guides", "geometry.adjustments", "geometry.connectionSites", "geometry.adjustmentHandles"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Image when source.Editable:
                output.Add(new("replaceImage", ["image.asset"]));
                if (!string.IsNullOrEmpty(element.Image.SvgAssetId))
                    output.Add(new("replaceSvg", ["image.svgAsset"]));
                output.Add(new("setImageCrop", ["image.crop"]));
                output.Add(new("setImageFit", ["image.fit"]));
                output.Add(new("setFrame", EditableFrameFields));
                output.Add(new("setOpacity", ["opacity"]));
                output.Add(new("setImageEffects", ["image.border", "image.shadow", "image.glow", "image.innerShadow", "image.reflection", "image.softEdge"]));
                if (element.Image.CustomMaskPaths.Count == 0 ||
                    CanProjectCustomGeometry(ImageMaskShape(element.Image)))
                {
                    // Keep one capability per picture mask owner.  The
                    // compiler still gates each transition by topology and
                    // profile, while the field set covers preset identity,
                    // complete adjustment lists, and literal path changes.
                    output.Add(new("setImageMask", ["image.mask.preset", "image.mask.adjustments", "image.mask.paths"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Chart when source.Editable:
                output.Add(new("setChartTitle", ["chart.title"]));
                output.Add(new("setChartData", ["chart.data"]));
                output.Add(new("setChartTextStyle", ["chart.textStyle", "chart.fontFamily"]));
                output.Add(new("setChartFill", ["chart.fill", "chart.legendFill"]));
                // Direct series stroke and marker graphs are parsed by the
                // existing ChartSpace codecs.  Keep this capability separate
                // from fill/data so a source-bound edit can change only the
                // series style leaves and still fail closed by chart family
                // in the compiler (scatter uses marker.line, not series.line).
                output.Add(new("setChartSeriesStyle", [
                    "chart.data.series[].stroke",
                    "chart.data.series[].marker",
                    "chart.data.series[].explosion",
                    "chart.data.series[].valuesFormatCode",
                ]));
                // Trendlines and error bars are direct c:series children with
                // bounded readers/writers of their own.  Keep them separate
                // from paint. Trendlines own their ordered list, including
                // insertion/removal; local error bars own their optional
                // object. Unprojected formula sources remain protected.
                output.Add(new("setChartSeriesAnalytics", [
                    "chart.data.series[].trendlines",
                    "chart.data.series[].errorBars",
                ]));
                if (element.Chart.Frame is not null)
                    output.Add(new("setChartFrame", ["chart.frame"]));
                output.Add(new("setChartLabels", ["chart.labels"]));
                if (element.Chart.XAxis is not null || element.Chart.YAxis is not null ||
                    element.Chart.SecondaryXAxis is not null || element.Chart.SecondaryYAxis is not null)
                    output.Add(new("setChartAxis", ["chart.axis"]));
                // Common ChartSpace plot scalars (legend placement, grouping,
                // gap width, axis visibility, line smoothness/colour
                // variation, scatter style, and bounded circular/bubble geometry) are safe
                // only for the parser-owned chart families below.  The
                // capability is deliberately broad at the operation level;
                // the source-bound compiler still checks each field and the
                // chart family before writing the existing ChartPart.
                if (element.Chart.Type is SpreadsheetChartType.Bar or
                    SpreadsheetChartType.Line or
                    SpreadsheetChartType.Area or
                    SpreadsheetChartType.Pie or
                    SpreadsheetChartType.Doughnut or
                    SpreadsheetChartType.Bubble or
                    SpreadsheetChartType.Radar or
                    SpreadsheetChartType.Combo or
                    SpreadsheetChartType.Scatter)
                    output.Add(new("setChartPlot", ["chart.plot"]));
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Table when source.Editable:
                output.Add(new("replaceText", ["text"]));
                output.Add(new("setTableStyle", ["table.style"]));
                // The native rectangular table profile keeps grid identity
                // and merge topology fixed, but its existing column widths
                // and row heights are safe scalar leaves. Expose them as a
                // separate capability so a geometry edit cannot be confused
                // with text reflow or a topology change.
                output.Add(new("setTableGeometry", ["table.geometry"]));
                if (element.Table.CellStyleEditable || element.Table.CellTextStyleEditable)
                {
                    var fields = new List<string>();
                    if (element.Table.CellStyleEditable)
                    {
                        fields.Add("table.cell.fill");
                        fields.Add("table.cell.borders");
                    }
                    if (element.Table.CellTextStyleEditable)
                        fields.Add("table.cell.textStyle");
                    output.Add(new("setTableCellStyle", fields));
                }
                if (element.Table.Rows.Any(row => row.Cells.Any(cell =>
                        cell.TextBody is not null &&
                        PptxTableCodec.IsBoundedMixedRunTextBody(cell.TextBody) &&
                        cell.TextBody.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run =>
                            run.ContentCase == PresentationTextRun.ContentOneofCase.Field &&
                            Guid.TryParseExact(run.Field.Id, "B", out _) &&
                            PptxTextCodec.ValidFieldType(run.Field.Type) &&
                            !PptxTextCodec.IsAutomaticFieldType(run.Field.Type)))))
                {
                    var fields = new List<string> { "table.rows[].cells[].text.paragraphs[].runs[].field.type" };
                    if (element.Table.Rows.Any(row => row.Cells.Any(cell =>
                            cell.TextBody is not null &&
                            PptxTableCodec.IsBoundedMixedRunTextBody(cell.TextBody) &&
                            cell.TextBody.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run =>
                                run.ContentCase == PresentationTextRun.ContentOneofCase.Field &&
                                Guid.TryParseExact(run.Field.Id, "B", out _)))))
                        fields.Add("table.rows[].cells[].text.paragraphs[].runs[].field.id");
                    output.Add(new("setTextField", fields));
                }
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Connector when source.Editable:
                output.Add(new("setStroke", ["stroke"]));
                output.Add(new("setConnectorArrows", ["startArrow", "endArrow", "startArrowWidth", "startArrowLength", "endArrowWidth", "endArrowLength"]));
                output.Add(new("setConnectorType", ["connectorType", "bendAdjustment"]));
                if (customSiteEndpoints || element.Connector.StartTargetId.Length == 0 && element.Connector.EndTargetId.Length == 0)
                    output.Add(new("setConnectorEndpoints", ["from", "to"]));
                if (!customSiteEndpoints && element.Connector.StartFrameAnchor is null && element.Connector.EndFrameAnchor is null)
                    output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Group when source.Editable:
                output.Add(new("setFrame", EditableFrameFields));
                break;
            case PresentationElement.ContentOneofCase.Diagram:
                output.Add(new("setSmartArtText", ["smartArt.text"]));
                output.Add(new("setSmartArtGraph", ["smartArt.connections"]));
                if (HasSmartArtPicturePaintProfile(element.Diagram))
                {
                    output.Add(new("setSmartArtImage", ["smartArt.nodes[].asset"]));
                    output.Add(new("setSmartArtImagePaint", ["smartArt.nodes[].image"]));
                }
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                if (element.Diagram.DrawingCacheVerified)
                    output.Add(new("detachSmartArt", ["smartArt.detachToShapes"]));
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (element.Opaque.DiagramText is { Nodes.Count: > 0 })
                    output.Add(new("setSmartArtText", ["smartArt.text"]));
                if (element.Opaque.OleWorkbook is not null || element.Opaque.OleOfficePackage is not null)
                    output.Add(new("setOlePayload", ["ole.payload"]));
                if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
        }
        if (source.VisibilityEditable) output.Add(new("setHidden", ["hidden"]));
        if (source.LockingEditable) output.Add(new("setLocked", ["locked"]));
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static bool HasSmartArtPicturePaintProfile(PresentationDiagram diagram)
    {
        if (!diagram.DrawingCacheVerified || diagram.Nodes.Count == 0 ||
            diagram.Nodes.Any(node => string.IsNullOrWhiteSpace(node.AssetId)))
            return false;
        return diagram.Nodes.All(node => diagram.Drawing?.Children
            .SingleOrDefault(child => child.Name == node.Id && child.Shape is not null)
            ?.Shape?.ImageFill is { AssetId.Length: > 0 });
    }

    private static bool TextTopologyRepresentable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase is PresentationTextRun.ContentOneofCase.Text or PresentationTextRun.ContentOneofCase.LineBreak or PresentationTextRun.ContentOneofCase.Field));

    private static IReadOnlyList<CapabilitySpec> OpaqueCapabilities(PresentationElement element)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        if (element.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
            PpjNativeTextProjection.TryRead(element.Opaque.RawXml, out _))
            output.Add(new("replaceText", ["visibleText"]));
        if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
        if (source.VisibilityEditable) output.Add(new("setHidden", ["hidden"]));
        if (source.LockingEditable) output.Add(new("setLocked", ["locked"]));
        if (source.AccessibilityEditable) output.Add(new("setAccessibility", ["accessibility"]));
        // OLE and source-owned chart payloads stay opaque until the
        // corresponding PPJ typed state is projected. Proven diagram text is
        // handled by ProjectSourceSmartArt before this fallback is selected.
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static JsonObject NativeRef(
        ProjectionContext context,
        string scope,
        string objectHash,
        IEnumerable<CapabilitySpec> capabilities,
        JsonArray? leaves = null)
    {
        var capabilityArray = new JsonArray();
        foreach (var capability in capabilities
                     .OrderBy(item => item.Operation, StringComparer.Ordinal)
                     .ThenBy(item => string.Join("\0", item.Fields), StringComparer.Ordinal))
        {
            var fields = new JsonArray();
            foreach (var field in capability.Fields) fields.Add(StringNode(field));
            capabilityArray.Add(new JsonObject
            {
                ["id"] = StringNode($"cap-{capability.Operation}-{Sha256(Encoding.UTF8.GetBytes(scope + capability.Operation))[..10]}"),
                ["operation"] = StringNode(capability.Operation),
                ["expectedHash"] = StringNode(objectHash),
                ["fields"] = fields,
            });
        }
        var output = new JsonObject
        {
            ["handle"] = StringNode($"nr-{Sha256(Encoding.UTF8.GetBytes(context.SourceSha256 + "\0" + scope))}"),
            ["sourceSha256"] = StringNode(context.SourceSha256),
            ["revision"] = StringNode(context.Revision),
            ["objectHash"] = StringNode(objectHash),
            ["capabilitySetSha256"] = StringNode(Sha256(CanonicalBytes(capabilityArray))),
            ["capabilities"] = capabilityArray,
        };
        if (leaves is { Count: > 0 }) output["leaves"] = leaves;
        return output;
    }

    private static JsonObject ShapeFrame(PresentationShape shape)
    {
        var frame = Frame(shape.LeftEmu, shape.TopEmu, shape.WidthEmu, shape.HeightEmu);
        if (shape.Transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(shape.Transform.RotationAngle60000 / 60_000d);
        if (shape.Transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(shape.Transform.FlipHorizontal);
        if (shape.Transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(shape.Transform.FlipVertical);
        return frame;
    }

    private static JsonObject ShapeFrame(PresentationPlaceholderFrame source)
    {
        var frame = Frame(source.LeftEmu, source.TopEmu, source.WidthEmu, source.HeightEmu);
        if (source.HasRotationAngle60000)
            frame["rotation"] = JsonValue.Create(source.RotationAngle60000 / 60_000d);
        if (source.HasFlipHorizontal) frame["flipH"] = JsonValue.Create(source.FlipHorizontal);
        if (source.HasFlipVertical) frame["flipV"] = JsonValue.Create(source.FlipVertical);
        return frame;
    }

    private static readonly string[] EditableFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height", "frame.rotation", "frame.flipH", "frame.flipV"];

    private static readonly string[] PositionFrameFields =
        ["frame.x", "frame.y", "frame.width", "frame.height"];

    private static JsonObject ImageFrame(PresentationImage image)
    {
        var frame = Frame(image.LeftEmu, image.TopEmu, image.WidthEmu, image.HeightEmu);
        if (image.Transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(image.Transform.RotationAngle60000 / 60_000d);
        if (image.Transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(image.Transform.FlipHorizontal);
        if (image.Transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(image.Transform.FlipVertical);
        return frame;
    }

    private static JsonObject TableFrame(PresentationTable table) =>
        Frame(table.LeftEmu, table.TopEmu, table.WidthEmu, table.HeightEmu, table.FrameTransform);

    private static JsonObject ChartFrame(PresentationChart chart) =>
        Frame(chart.LeftEmu, chart.TopEmu, chart.WidthEmu, chart.HeightEmu, chart.FrameTransform);

    private static JsonObject GroupFrame(PresentationGroup group) =>
        Frame(group.LeftEmu, group.TopEmu, group.WidthEmu, group.HeightEmu, group.FrameTransform);

    private static JsonObject DiagramFrame(PresentationDiagram diagram) =>
        Frame(diagram.LeftEmu, diagram.TopEmu, diagram.WidthEmu, diagram.HeightEmu);

    private static JsonObject ConnectorFrame(PresentationConnector connector)
    {
        var left = Math.Min(connector.StartXEmu, connector.EndXEmu);
        var top = Math.Min(connector.StartYEmu, connector.EndYEmu);
        return Frame(left, top, Math.Abs(connector.EndXEmu - connector.StartXEmu), Math.Abs(connector.EndYEmu - connector.StartYEmu));
    }

    private static JsonObject ElementFrame(PresentationElement element) => element.ContentCase switch
    {
        PresentationElement.ContentOneofCase.Shape => ShapeFrame(element.Shape),
        PresentationElement.ContentOneofCase.Image => ImageFrame(element.Image),
        PresentationElement.ContentOneofCase.Table => TableFrame(element.Table),
        PresentationElement.ContentOneofCase.Connector => ConnectorFrame(element.Connector),
        PresentationElement.ContentOneofCase.Chart => ChartFrame(element.Chart),
        PresentationElement.ContentOneofCase.Diagram => DiagramFrame(element.Diagram),
        PresentationElement.ContentOneofCase.Group => GroupFrame(element.Group),
        PresentationElement.ContentOneofCase.Opaque => Frame(element.Opaque.LeftEmu, element.Opaque.TopEmu, element.Opaque.WidthEmu, element.Opaque.HeightEmu),
        _ => Frame(0, 0, 1, 1),
    };

    private static JsonObject Frame(long left, long top, long width, long height) => new()
    {
        ["x"] = JsonValue.Create(Points(left)),
        ["y"] = JsonValue.Create(Points(top)),
        ["width"] = JsonValue.Create(Math.Max(0.001, Points(width))),
        ["height"] = JsonValue.Create(Math.Max(0.001, Points(height))),
    };

    private static JsonObject Frame(
        long left,
        long top,
        long width,
        long height,
        PresentationFrameTransform? transform)
    {
        var frame = Frame(left, top, width, height);
        if (transform?.HasRotationAngle60000 == true) frame["rotation"] = JsonValue.Create(transform.RotationAngle60000 / 60_000d);
        if (transform?.HasFlipHorizontal == true) frame["flipH"] = JsonValue.Create(transform.FlipHorizontal);
        if (transform?.HasFlipVertical == true) frame["flipV"] = JsonValue.Create(transform.FlipVertical);
        return frame;
    }

    private static JsonObject ConnectorEndpoint(
        string targetId,
        long x,
        long y,
        string pageId,
        ProjectionContext context,
        PresentationConnectorFrameAnchor? frameAnchor = null,
        uint? connectionSite = null)
    {
        if (frameAnchor is not null)
        {
            if (!context.TryElementId(pageId, frameAnchor.TargetId, out var anchorTarget))
                throw new CodecException("ppj.connector.endpoint", "Frame-anchor target cannot be mapped to the projected page.");
            return new JsonObject { ["element"] = StringNode(anchorTarget), ["anchor"] = StringNode(frameAnchor.Anchor) };
        }
        if (!string.IsNullOrEmpty(targetId) && context.TryElementId(pageId, targetId, out var projected))
            return connectionSite is { } index
                ? new JsonObject { ["element"] = StringNode(projected), ["connectionSite"] = JsonValue.Create(index) }
                : new JsonObject { ["element"] = StringNode(projected), ["anchor"] = StringNode("auto") };
        return new JsonObject { ["x"] = JsonValue.Create(Points(x)), ["y"] = JsonValue.Create(Points(y)) };
    }

    private static JsonObject Stroke(
        string rgb,
        long widthEmu,
        string? style,
        string? cap,
        string? join,
        double? opacity,
        string? scheme = null)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(scheme)
                ? new JsonObject { ["token"] = StringNode(scheme) }
                : StringNode(Color(string.IsNullOrEmpty(rgb) ? "000000" : rgb)),
                ["width"] = JsonValue.Create(Math.Max(0, Points(widthEmu))),
        };
        var dash = Dash(style);
        if (dash is not null) output["dash"] = StringNode(dash);
        if (cap is "flat" or "round" or "square") output["cap"] = StringNode(cap);
        if (join is "miter" or "round" or "bevel") output["join"] = StringNode(join);
        if (opacity is not null) output["opacity"] = JsonValue.Create(opacity.Value);
        return output;
    }

    private static JsonObject ChartTextReflection(PresentationReflection reflection)
    {
        var output = Reflection(reflection);
        if (reflection.HasFadeDirectionAngle60000) output["fadeAngle"] = JsonValue.Create(reflection.FadeDirectionAngle60000 / 60000d);
        if (reflection.HasScaleXThousandthPercent) output["scaleX"] = JsonValue.Create(reflection.ScaleXThousandthPercent / 100000d);
        if (reflection.HasScaleYThousandthPercent) output["scaleY"] = JsonValue.Create(reflection.ScaleYThousandthPercent / 100000d);
        if (reflection.HasSkewXAngle60000) output["skewX"] = JsonValue.Create(reflection.SkewXAngle60000 / 60000d);
        if (reflection.HasSkewYAngle60000) output["skewY"] = JsonValue.Create(reflection.SkewYAngle60000 / 60000d);
        if (reflection.HasAlignment) output["alignment"] = StringNode(reflection.Alignment);
        if (reflection.HasRotateWithShape) output["rotateWithShape"] = JsonValue.Create(reflection.RotateWithShape);
        if (reflection.HasStartPositionThousandthPercent) output["startPosition"] = JsonValue.Create(Unit(reflection.StartPositionThousandthPercent));
        if (reflection.HasEndPositionThousandthPercent) output["endPosition"] = JsonValue.Create(Unit(reflection.EndPositionThousandthPercent));
        if (!reflection.HasBlurRadiusEmu) output.Remove("blur");
        if (!reflection.HasDistanceEmu) output.Remove("distance");
        if (!reflection.HasDirectionAngle60000) output.Remove("angle");
        if (!reflection.HasStartOpacityThousandthPercent) output.Remove("startOpacity");
        if (!reflection.HasEndOpacityThousandthPercent) output.Remove("endOpacity");
        return output;
    }

    private static JsonObject ChartTextInnerShadow(PresentationInnerShadow shadow)
    {
        var output = InnerShadow(shadow);
        if (!shadow.HasBlurRadiusEmu) output.Remove("blur");
        if (!shadow.HasDistanceEmu) output.Remove("distance");
        if (!shadow.HasDirectionAngle60000) output.Remove("angle");
        return output;
    }

    private static JsonObject ChartTextShadow(PresentationShadow shadow)
    {
        var output = Shadow(shadow, includeOpacity: shadow.HasOpacityThousandthPercent);
        if (shadow.HasScaleXThousandthPercent) output["scaleX"] = JsonValue.Create(shadow.ScaleXThousandthPercent / 100000d);
        if (shadow.HasScaleYThousandthPercent) output["scaleY"] = JsonValue.Create(shadow.ScaleYThousandthPercent / 100000d);
        if (shadow.HasSkewXAngle60000) output["skewX"] = JsonValue.Create(shadow.SkewXAngle60000 / 60000d);
        if (shadow.HasSkewYAngle60000) output["skewY"] = JsonValue.Create(shadow.SkewYAngle60000 / 60000d);
        if (!shadow.HasBlurRadiusEmu) output.Remove("blur");
        if (!shadow.HasDistanceEmu) output.Remove("distance");
        if (!shadow.HasDirectionAngle60000) output.Remove("angle");
        return output;
    }

    private static JsonObject Shadow(PresentationShadow shadow, bool includeOpacity = true)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(shadow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(shadow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(shadow.ColorRgb) ? "000000" : shadow.ColorRgb)),
            ["blur"] = JsonValue.Create(Math.Max(0, Points(shadow.HasBlurRadiusEmu ? shadow.BlurRadiusEmu : 0))),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(shadow.HasDistanceEmu ? shadow.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((shadow.HasDirectionAngle60000 ? shadow.DirectionAngle60000 : 0) / 60_000d),
        };
        if (includeOpacity)
            output["opacity"] = JsonValue.Create(shadow.HasOpacityThousandthPercent ? Unit(shadow.OpacityThousandthPercent) : 1);
        if (shadow.HasAlignment) output["alignment"] = StringNode(shadow.Alignment);
        if (shadow.HasRotateWithShape) output["rotateWithShape"] = JsonValue.Create(shadow.RotateWithShape);
        return output;
    }

    private static JsonObject TextRunShadow(PresentationShadow shadow)
    {
        var output = Shadow(shadow);
        if (shadow.HasScaleXThousandthPercent)
            output["scaleX"] = JsonValue.Create(shadow.ScaleXThousandthPercent / 100000d);
        if (shadow.HasScaleYThousandthPercent)
            output["scaleY"] = JsonValue.Create(shadow.ScaleYThousandthPercent / 100000d);
        if (shadow.HasSkewXAngle60000)
            output["skewX"] = JsonValue.Create(shadow.SkewXAngle60000 / 60000d);
        if (shadow.HasSkewYAngle60000)
            output["skewY"] = JsonValue.Create(shadow.SkewYAngle60000 / 60000d);
        return output;
    }

    private static JsonObject Glow(PresentationGlow glow)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(glow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(glow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(glow.ColorRgb) ? "000000" : glow.ColorRgb)),
            ["radius"] = JsonValue.Create(Math.Max(0, Points(glow.HasRadiusEmu ? glow.RadiusEmu : 0))),
        };
        if (glow.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(glow.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject InnerShadow(PresentationInnerShadow shadow)
    {
        var output = new JsonObject
        {
            ["color"] = !string.IsNullOrEmpty(shadow.ColorScheme)
                ? new JsonObject { ["token"] = StringNode(shadow.ColorScheme) }
                : StringNode(Color(string.IsNullOrEmpty(shadow.ColorRgb) ? "000000" : shadow.ColorRgb)),
            ["blur"] = JsonValue.Create(Math.Max(0, Points(shadow.HasBlurRadiusEmu ? shadow.BlurRadiusEmu : 0))),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(shadow.HasDistanceEmu ? shadow.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((shadow.HasDirectionAngle60000 ? shadow.DirectionAngle60000 : 0) / 60_000d),
        };
        if (shadow.HasOpacityThousandthPercent)
            output["opacity"] = JsonValue.Create(Unit(shadow.OpacityThousandthPercent));
        return output;
    }

    private static JsonObject Reflection(
        PresentationReflection reflection,
        bool includePositions = false,
        bool includeFadeAngle = false,
        bool includeScaleX = false,
        bool includeScaleY = false,
        bool includeSkewX = false,
        bool includeSkewY = false,
        bool includeAlignment = false,
        bool includeRotateWithShape = false)
    {
        var output = new JsonObject
        {
            ["blur"] = JsonValue.Create(Math.Max(0, Points(reflection.HasBlurRadiusEmu ? reflection.BlurRadiusEmu : 0))),
            ["startOpacity"] = JsonValue.Create(Unit(reflection.HasStartOpacityThousandthPercent ? reflection.StartOpacityThousandthPercent : 100_000)),
            ["endOpacity"] = JsonValue.Create(Unit(reflection.HasEndOpacityThousandthPercent ? reflection.EndOpacityThousandthPercent : 0)),
            ["distance"] = JsonValue.Create(Math.Max(0, Points(reflection.HasDistanceEmu ? reflection.DistanceEmu : 0))),
            ["angle"] = JsonValue.Create((reflection.HasDirectionAngle60000 ? reflection.DirectionAngle60000 : 0) / 60_000d),
        };
        if (includeFadeAngle && reflection.HasFadeDirectionAngle60000)
            output["fadeAngle"] = JsonValue.Create(reflection.FadeDirectionAngle60000 / 60_000d);
        if (includeScaleX && reflection.HasScaleXThousandthPercent)
            output["scaleX"] = JsonValue.Create(reflection.ScaleXThousandthPercent / 100_000d);
        if (includeScaleY && reflection.HasScaleYThousandthPercent)
            output["scaleY"] = JsonValue.Create(reflection.ScaleYThousandthPercent / 100_000d);
        if (includeSkewX && reflection.HasSkewXAngle60000)
            output["skewX"] = JsonValue.Create(reflection.SkewXAngle60000 / 60_000d);
        if (includeSkewY && reflection.HasSkewYAngle60000)
            output["skewY"] = JsonValue.Create(reflection.SkewYAngle60000 / 60_000d);
        if (includeAlignment && reflection.HasAlignment)
            output["alignment"] = StringNode(reflection.Alignment);
        if (includeRotateWithShape && reflection.HasRotateWithShape)
            output["rotateWithShape"] = JsonValue.Create(reflection.RotateWithShape);
        if (includePositions && reflection.HasStartPositionThousandthPercent)
            output["startPosition"] = JsonValue.Create(Unit(reflection.StartPositionThousandthPercent));
        if (includePositions && reflection.HasEndPositionThousandthPercent)
            output["endPosition"] = JsonValue.Create(Unit(reflection.EndPositionThousandthPercent));
        return output;
    }

    private static JsonObject SoftEdge(PresentationSoftEdge softEdge) => new()
    {
        ["radius"] = JsonValue.Create(Math.Max(0, Points(softEdge.HasRadiusEmu ? softEdge.RadiusEmu : 0))),
    };

    private static JsonObject? Accessibility(PresentationNonVisualAccessibility? value)
    {
        if (value is null) return null;
        var output = new JsonObject { ["decorative"] = JsonValue.Create(value.HasDecorative && value.Decorative) };
        if (!string.IsNullOrEmpty(value.Title)) output["title"] = StringNode(value.Title);
        if (!string.IsNullOrEmpty(value.Description)) output["description"] = StringNode(value.Description);
        return output;
    }

    private static JsonObject? ImageAccessibility(PresentationImage image)
    {
        if (!image.HasAccessibilityDecorative && string.IsNullOrEmpty(image.AccessibilityTitle) && string.IsNullOrEmpty(image.AltText)) return null;
        var output = new JsonObject { ["decorative"] = JsonValue.Create(image.HasAccessibilityDecorative && image.AccessibilityDecorative) };
        if (!string.IsNullOrEmpty(image.AccessibilityTitle)) output["title"] = StringNode(image.AccessibilityTitle);
        if (!string.IsNullOrEmpty(image.AltText)) output["description"] = StringNode(image.AltText);
        return output;
    }

    private static string PlaceholderType(string value) => value switch
    {
        "ctrTitle" or "title" => "title",
        "subTitle" or "subtitle" => "subtitle",
        "body" => "body",
        "obj" or "content" => "content",
        "pic" or "picture" => "picture",
        "chart" => "chart",
        "tbl" or "table" => "table",
        "dt" or "date" => "date",
        "ftr" or "footer" => "footer",
        "sldNum" or "slide-number" => "slide-number",
        _ => "other",
    };

    private static string? ChartType(SpreadsheetChartType type) => type switch
    {
        SpreadsheetChartType.Bar => "column",
        SpreadsheetChartType.Line => "line",
        SpreadsheetChartType.Pie => "pie",
        SpreadsheetChartType.Area => "area",
        SpreadsheetChartType.Doughnut => "doughnut",
        SpreadsheetChartType.Scatter => "scatter",
        SpreadsheetChartType.Bubble => "bubble",
        SpreadsheetChartType.Radar => "radar",
        SpreadsheetChartType.Combo => "combo",
        _ => null,
    };

    private static string? Marker(SpreadsheetChartMarkerSymbol value) => value switch
    {
        SpreadsheetChartMarkerSymbol.None => "none",
        SpreadsheetChartMarkerSymbol.Dot => "dot",
        SpreadsheetChartMarkerSymbol.Circle => "circle",
        SpreadsheetChartMarkerSymbol.Square => "square",
        SpreadsheetChartMarkerSymbol.Diamond => "diamond",
        SpreadsheetChartMarkerSymbol.Triangle => "triangle",
        SpreadsheetChartMarkerSymbol.X => "x",
        SpreadsheetChartMarkerSymbol.Star => "star",
        SpreadsheetChartMarkerSymbol.Plus => "plus",
        SpreadsheetChartMarkerSymbol.Dash => "dash",
        _ => null,
    };

    private static string? ChartDash(SpreadsheetChartLineDashStyle value) => value switch
    {
        SpreadsheetChartLineDashStyle.Solid => "solid",
        SpreadsheetChartLineDashStyle.Dashed => "dash",
        SpreadsheetChartLineDashStyle.Dotted => "dot",
        SpreadsheetChartLineDashStyle.DashDot => "dash-dot",
        _ => null,
    };

    private static string? DataLabelPosition(SpreadsheetChartDataLabelPosition value) => value switch
    {
        SpreadsheetChartDataLabelPosition.BestFit => "best-fit",
        SpreadsheetChartDataLabelPosition.Bottom => "bottom",
        SpreadsheetChartDataLabelPosition.Center => "center",
        SpreadsheetChartDataLabelPosition.InsideBase => "inside-base",
        SpreadsheetChartDataLabelPosition.InsideEnd => "inside-end",
        SpreadsheetChartDataLabelPosition.Left => "left",
        SpreadsheetChartDataLabelPosition.OutsideEnd => "outside-end",
        SpreadsheetChartDataLabelPosition.Right => "right",
        SpreadsheetChartDataLabelPosition.Top => "top",
        _ => null,
    };

    private static string? TrendlineType(SpreadsheetChartTrendlineType value) => value switch
    {
        SpreadsheetChartTrendlineType.Exponential => "exponential",
        SpreadsheetChartTrendlineType.Linear => "linear",
        SpreadsheetChartTrendlineType.Logarithmic => "logarithmic",
        SpreadsheetChartTrendlineType.MovingAverage => "moving-average",
        SpreadsheetChartTrendlineType.Polynomial => "polynomial",
        SpreadsheetChartTrendlineType.Power => "power",
        _ => null,
    };

    private static string? ErrorBarDirection(SpreadsheetChartErrorBarDirection value) => value switch
    {
        SpreadsheetChartErrorBarDirection.X => "x",
        SpreadsheetChartErrorBarDirection.Y => "y",
        _ => null,
    };

    private static string? ErrorBarType(SpreadsheetChartErrorBarType value) => value switch
    {
        SpreadsheetChartErrorBarType.Both => "both",
        SpreadsheetChartErrorBarType.Minus => "minus",
        SpreadsheetChartErrorBarType.Plus => "plus",
        _ => null,
    };

    private static string? ErrorBarValueType(SpreadsheetChartErrorBarValueType value) => value switch
    {
        SpreadsheetChartErrorBarValueType.Custom => "custom",
        SpreadsheetChartErrorBarValueType.FixedValue => "fixed-value",
        SpreadsheetChartErrorBarValueType.Percentage => "percentage",
        SpreadsheetChartErrorBarValueType.StandardDeviation => "standard-deviation",
        SpreadsheetChartErrorBarValueType.StandardError => "standard-error",
        _ => null,
    };

    private static string? Dash(string? value) => value switch
    {
        null or "" or "solid" => "solid",
        "dashed" or "dash" => "dash",
        "dotted" or "dot" => "dot",
        "dash-dot" or "dashDot" => "dash-dot",
        "long-dash" or "longDash" => "long-dash",
        _ => null,
    };

    private static string? Arrow(string? value) => value switch
    {
        "triangle" or "stealth" or "diamond" or "oval" or "open" => value,
        "arrow" => "open",
        _ => null,
    };

    private static string Color(string rgb) => $"#{rgb.TrimStart('#').ToUpperInvariant()}";

    private static JsonNode BulletColor(string? rgb, string? scheme, bool hasOpacity, uint opacity)
    {
        if (!string.IsNullOrEmpty(rgb) && hasOpacity)
        {
            var alphaByte = Math.Clamp((int)Math.Round(Unit(opacity) * 255), 0, 255);
            if (Math.Round(alphaByte / 255d * 100_000) != opacity)
                return new JsonObject { ["rgb"] = StringNode(Color(rgb)), ["alpha"] = JsonValue.Create(Unit(opacity)) };
        }
        return TextColor(rgb, scheme, hasOpacity, opacity);
    }

    private static JsonNode TextColor(string? rgb, string? scheme, bool hasOpacity, uint opacity)
    {
        if (!string.IsNullOrEmpty(rgb))
        {
            var value = Color(rgb);
            if (!hasOpacity) return StringNode(value);
            var alpha = Math.Clamp((int)Math.Round(Unit(opacity) * 255), 0, 255);
            return StringNode($"{value}{alpha:X2}");
        }
        var output = new JsonObject { ["token"] = StringNode(scheme ?? string.Empty) };
        if (hasOpacity) output["alpha"] = JsonValue.Create(Unit(opacity));
        return output;
    }
    private static double Crop(int value) => Math.Clamp(value / 100_000d, -1, 1);
    private static double Unit(uint value) => Math.Clamp(value / 100_000d, 0, 1);
    private static double Points(long emu) => Math.Round(emu / EmuPerPoint, 6, MidpointRounding.AwayFromZero);

    private static string HashOrFallback(string? hash, IMessage fallback) =>
        IsCanonicalSha256(hash) ? hash! : Sha256(fallback.ToByteArray());

    private static string HashOrFallback(string? hash, ByteString fallback) =>
        IsCanonicalSha256(hash) ? hash! : Sha256(fallback.Span);

    private static bool IsCanonicalSha256(string? hash) =>
        hash is { Length: 64 } && hash.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

    private static string StableDocumentId(string? candidate, string sha256)
    {
        var normalized = NormalizeId(candidate, "presentation");
        return normalized.Length <= 100 ? normalized : $"presentation-{sha256[..24]}";
    }

    private static string NormalizeId(string? value, string fallback)
    {
        var normalized = InvalidIdCharacters().Replace(value ?? string.Empty, "-").Trim('-', '.', ':', '_');
        if (normalized.Length == 0 || !char.IsAsciiLetterOrDigit(normalized[0])) normalized = $"{fallback}-{normalized}".TrimEnd('-');
        if (normalized.Length > 112)
            normalized = $"{normalized[..87]}-{Sha256(Encoding.UTF8.GetBytes(value ?? fallback))[..24]}";
        return normalized;
    }

    private static byte[] CanonicalBytes(JsonNode node)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            node.WriteTo(writer);
        using var document = JsonDocument.Parse(buffer.WrittenMemory);
        return PpjCanonicalJson.Write(document.RootElement);
    }

    private static JsonNode StringNode(string value)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            writer.WriteStringValue(value);
        return JsonNode.Parse(buffer.WrittenSpan) ?? throw new InvalidOperationException("String JSON primitive could not be created.");
    }

    private static JsonNode NumberNode(double value)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            writer.WriteNumberValue(value);
        return JsonNode.Parse(buffer.WrittenSpan) ?? throw new InvalidOperationException("Number JSON primitive could not be created.");
    }

    private static string Sha256(byte[] bytes) => Sha256(bytes.AsSpan());
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    [GeneratedRegex("[^A-Za-z0-9._:-]+")]
    private static partial Regex InvalidIdCharacters();

    private sealed record CapabilitySpec(string Operation, IReadOnlyList<string> Fields);

    private sealed class ProjectionContext
    {
        private readonly IReadOnlyDictionary<string, Asset> sourceAssets;
        private readonly IReadOnlyDictionary<string, OpaqueOpcPart> sourceParts;
        private readonly PptxPackageSource sourcePackage;
        private readonly Dictionary<string, string> pageIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> masterIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> layoutIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> customShowIds = new(StringComparer.Ordinal);
        private readonly Dictionary<(string Page, string Element), string> elementIds = new();
        private readonly HashSet<string> usedIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdBySourceId = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdByHash = new(StringComparer.Ordinal);
        private readonly JsonArray programAssets = new();
        private readonly List<Asset> resultAssets = [];
        private readonly List<Asset> nativeSourceAssets = [];
        private readonly JsonArray nodes = new();
        private readonly Dictionary<string, PpjNativeLeafBinding> nativeLeafBindings = new(StringComparer.Ordinal);

        internal ProjectionContext(
            string sourceSha256,
            string revision,
            string assetRoot,
            IEnumerable<Asset> assets,
            OpaqueOpcGraph? opaque,
            PptxPackageSource sourcePackage)
        {
            SourceSha256 = sourceSha256;
            Revision = revision;
            AssetRoot = assetRoot;
            sourceAssets = assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
            sourceParts = (opaque?.Parts ?? []).ToDictionary(part => part.Path, StringComparer.OrdinalIgnoreCase);
            this.sourcePackage = sourcePackage;
        }

        internal string SourceSha256 { get; }
        internal string Revision { get; }
        internal IReadOnlyDictionary<string, PpjNativeLeafBinding> NativeLeafBindings => nativeLeafBindings;

        internal void RecordNativeLeaf(PpjNativeLeafBinding binding)
        {
            if (!nativeLeafBindings.TryAdd(binding.Id, binding))
                throw new CodecException("ppj.nativeRef.leafId", $"Duplicate projected native leaf ID {binding.Id}.");
        }
        internal string AssetRoot { get; }
        internal int VisibleObjectCount { get; private set; }
        internal JsonArray ProgramAssets => programAssets;
        internal IReadOnlyList<Asset> ResultAssets => resultAssets;
        internal IReadOnlyList<Asset> NativeSourceAssets => nativeSourceAssets;

        internal void ReleaseProjectionJson()
        {
            programAssets.Clear();
            nodes.Clear();
        }

        internal string RegisterPage(string sourceId, string? stableSourceId)
        {
            var id = UniqueId($"page-{NormalizeId(stableSourceId, NormalizeId(sourceId, "slide"))}");
            pageIds[sourceId] = id;
            return id;
        }

        internal string RegisterMaster(string sourceId)
        {
            var id = UniqueId($"master-{NormalizeId(sourceId, "master")}");
            masterIds[sourceId] = id;
            return id;
        }

        internal string RegisterLayout(string sourceId)
        {
            var id = UniqueId($"layout-{NormalizeId(sourceId, "layout")}");
            layoutIds[sourceId] = id;
            return id;
        }

        internal string RegisterCustomShow(string sourceId)
        {
            var id = UniqueId($"show-{NormalizeId(sourceId, "show")}");
            customShowIds[sourceId] = id;
            return id;
        }

        internal string RegisterElement(string pageId, string sourceId)
        {
            var id = UniqueId($"{pageId}-{NormalizeId(PageLocalElementPath(sourceId), "element")}");
            elementIds[(pageId, sourceId)] = id;
            return id;
        }

        private static string PageLocalElementPath(string sourceId)
        {
            const string marker = "/element/";
            var index = sourceId.IndexOf(marker, StringComparison.Ordinal);
            return index < 0 ? sourceId : sourceId[(index + 1)..];
        }

        internal string PageId(string sourceId) => pageIds[sourceId];
        internal bool TryPageId(string sourceId, out string id) => pageIds.TryGetValue(sourceId, out id!);
        internal string MasterId(string sourceId) => masterIds[sourceId];
        internal bool TryLayoutId(string sourceId, out string id) => layoutIds.TryGetValue(sourceId, out id!);
        internal string LayoutId(string sourceId) => layoutIds[sourceId];
        internal string CustomShowId(string sourceId) => customShowIds[sourceId];
        internal bool TryCustomShowId(string sourceId, out string id) => customShowIds.TryGetValue(sourceId, out id!);
        internal string ElementId(string pageId, string sourceId) => elementIds[(pageId, sourceId)];
        internal bool TryElementId(string pageId, string sourceId, out string id) => elementIds.TryGetValue((pageId, sourceId), out id!);

        internal string UniqueId(string candidate)
        {
            var normalized = NormalizeId(candidate, "id");
            if (usedIds.Add(normalized)) return normalized;
            var suffix = Sha256(Encoding.UTF8.GetBytes(candidate))[..12];
            var shortened = normalized.Length > 114 ? normalized[..114] : normalized;
            var unique = $"{shortened}-{suffix}";
            var ordinal = 2;
            while (!usedIds.Add(unique)) unique = $"{shortened}-{suffix}-{ordinal++}";
            return unique;
        }

        internal bool TryMaterializeAsset(string sourceId, out string programAssetId)
        {
            if (assetIdBySourceId.TryGetValue(sourceId, out programAssetId!)) return true;
            if (!sourceAssets.TryGetValue(sourceId, out var source)) return false;
            var hash = HashOrFallback(source.Sha256, source.Data);
            if (assetIdByHash.TryGetValue(hash, out programAssetId!))
            {
                assetIdBySourceId[sourceId] = programAssetId;
                return true;
            }
            programAssetId = UniqueId($"asset-{hash[..20]}");
            assetIdBySourceId[sourceId] = programAssetId;
            assetIdByHash[hash] = programAssetId;
            var extension = Extension(source.ContentType, source.FileName);
            var fileName = $"{hash}.{extension}";
            programAssets.Add(new JsonObject
            {
                ["id"] = StringNode(programAssetId),
                ["uri"] = StringNode($"{AssetRoot}/{fileName}"),
                ["mimeType"] = StringNode(source.ContentType),
                ["sha256"] = StringNode(hash),
                ["rights"] = new JsonObject { ["status"] = StringNode("user-provided") },
                ["accessibility"] = new JsonObject { ["decorative"] = JsonValue.Create(false), ["description"] = StringNode("Imported source media.") },
            });
            var materialized = source.Clone();
            materialized.Id = programAssetId;
            materialized.FileName = fileName;
            materialized.Sha256 = hash;
            resultAssets.Add(materialized);
            return true;
        }

        internal bool TryMaterializeSourcePart(string partPath, string contentType, out string programAssetId)
        {
            var sourceId = $"source-part:{partPath}";
            if (assetIdBySourceId.TryGetValue(sourceId, out programAssetId!)) return true;
            if (!sourceParts.TryGetValue(partPath, out var metadata) ||
                !metadata.ContentType.Equals(contentType, StringComparison.OrdinalIgnoreCase))
            {
                programAssetId = string.Empty;
                return false;
            }

            byte[] data;
            using (var stream = sourcePackage.OpenRead())
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
            {
                var entry = archive.Entries.SingleOrDefault(candidate => candidate.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));
                if (entry is null || entry.Length is <= 0 or > 16 * 1024 * 1024)
                {
                    programAssetId = string.Empty;
                    return false;
                }
                using var entryStream = entry.Open();
                using var memory = new MemoryStream();
                entryStream.CopyTo(memory);
                data = memory.ToArray();
            }
            var hash = Sha256(data);
            if (!hash.Equals(metadata.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                programAssetId = string.Empty;
                return false;
            }

            if (!assetIdByHash.TryGetValue(hash, out programAssetId!))
            {
                programAssetId = UniqueId($"asset-{hash[..20]}");
                assetIdByHash[hash] = programAssetId;
                var extension = Extension(contentType, Path.GetFileName(partPath));
                var fileName = $"{hash}.{extension}";
                programAssets.Add(new JsonObject
                {
                    ["id"] = StringNode(programAssetId),
                    ["uri"] = StringNode($"{AssetRoot}/{fileName}"),
                    ["mimeType"] = StringNode(contentType),
                    ["sha256"] = StringNode(hash),
                    ["rights"] = new JsonObject { ["status"] = StringNode("user-provided") },
                    ["accessibility"] = new JsonObject { ["decorative"] = JsonValue.Create(false), ["description"] = StringNode("Imported embedded Office package.") },
                });
                resultAssets.Add(new Asset
                {
                    Id = programAssetId,
                    FileName = fileName,
                    ContentType = contentType,
                    Data = ByteString.CopyFrom(data),
                    Sha256 = hash,
                });
            }
            assetIdBySourceId[sourceId] = programAssetId;
            var nativeId = PptxAssetCatalog.NativeAssetIdFor(contentType, hash);
            if (!nativeSourceAssets.Any(asset => asset.Id.Equals(nativeId, StringComparison.Ordinal)))
                nativeSourceAssets.Add(new Asset
                {
                    Id = nativeId,
                    FileName = Path.GetFileName(partPath),
                    ContentType = contentType,
                    Data = ByteString.CopyFrom(data),
                    Sha256 = hash,
                });
            return true;
        }

        internal JsonObject? SmartArtNativeSections(IEnumerable<string> partPaths)
        {
            var sections = new SortedDictionary<string, string>(StringComparer.Ordinal);
            foreach (var partPath in partPaths)
            {
                if (!sourceParts.TryGetValue(partPath, out var part) || string.IsNullOrWhiteSpace(part.Sha256)) continue;
                var name = part.ContentType switch
                {
                    var value when value.Contains("diagramData", StringComparison.OrdinalIgnoreCase) => "dataSha256",
                    var value when value.Contains("diagramLayout", StringComparison.OrdinalIgnoreCase) => "layoutSha256",
                    var value when value.Contains("diagramStyle", StringComparison.OrdinalIgnoreCase) => "styleSha256",
                    var value when value.Contains("diagramColors", StringComparison.OrdinalIgnoreCase) => "colorsSha256",
                    var value when value.Contains("diagramDrawing", StringComparison.OrdinalIgnoreCase) => "drawingSha256",
                    _ => string.Empty,
                };
                if (name.Length > 0) sections.TryAdd(name, part.Sha256.ToLowerInvariant());
            }
            if (sections.Count == 0) return null;
            var output = new JsonObject();
            foreach (var (name, hash) in sections) output[name] = StringNode(hash);
            output["closureSha256"] = StringNode(Sha256(Encoding.UTF8.GetBytes(string.Join(
                "\n",
                sections.Select(pair => $"{pair.Key}:{pair.Value}")))));
            return output;
        }

        internal void RecordNode(string pageId, string id, string type, JsonObject nativeRef)
        {
            VisibleObjectCount++;
            nodes.Add(new JsonObject
            {
                ["id"] = StringNode(id),
                ["page"] = StringNode(pageId),
                ["type"] = StringNode(type),
                ["handle"] = StringNode(nativeRef["handle"]!.GetValue<string>()),
                ["objectHash"] = StringNode(nativeRef["objectHash"]!.GetValue<string>()),
                ["capabilitySetSha256"] = StringNode(nativeRef["capabilitySetSha256"]!.GetValue<string>()),
            });
        }

        internal JsonObject BuildNodeMap() => new()
        {
            ["schema"] = StringNode("office-kit/ppj-node-map/v1"),
            ["sourceSha256"] = StringNode(SourceSha256),
            ["revision"] = StringNode(Revision),
            ["nodes"] = nodes,
        };

        private static string Extension(string contentType, string fileName)
        {
            var fromType = contentType.ToLowerInvariant() switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/svg+xml" => "svg",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => "xlsx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => "docx",
                _ => string.Empty,
            };
            if (fromType.Length > 0) return fromType;
            var extension = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            return Regex.IsMatch(extension, "^[a-z0-9]{1,8}$") ? extension : "bin";
        }
    }
}
