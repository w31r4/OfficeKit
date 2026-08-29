using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

/// <summary>
/// Selects the authored or source-bound PPJ lowering path after one language
/// validation. Source-bound programs always compile against the exact PPTX
/// bytes named by their source descriptor; PPJ never turns a stale descriptor
/// into write authority.
/// </summary>
internal static class PpjPresentationCompiler
{
    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        byte[] sourceBytes,
        EffectiveCodecLimits limits)
    {
        var validation = PpjProgramValidator.Validate(request.ProgramJson.Span);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        if (validation.Program!.Source is null)
        {
            if (sourceBytes.Length != 0)
                throw new CodecException(
                    "ppj.unexpectedSource",
                    "A source-free PPJ compile cannot attach a PPTX source package.",
                    "$.source");
            return PpjAuthoredPresentationCompiler.Compile(request, limits);
        }

        return PpjSourceBoundPresentationCompiler.Compile(request, sourceBytes, limits, validation);
    }
}

/// <summary>
/// Reprojects an exact source package, compares it with the requested PPJ by
/// stable IDs, and maps only capability-issued semantic differences onto the
/// existing source-preserving Presentation IR. The mature PPTX writer remains
/// the independent package-level oracle; this class never writes OOXML.
/// </summary>
internal static class PpjSourceBoundPresentationCompiler
{
    private const double EmuPerPoint = 12_700d;

    internal static PpjCompileResult Compile(
        PresentationProgramRequest request,
        byte[] sourceBytes,
        EffectiveCodecLimits limits,
        PpjValidationResult validation)
    {
        if (sourceBytes.Length == 0)
            throw new CodecException(
                "ppj.source.missing",
                "A source-bound PPJ compile requires the exact source PPTX bytes.",
                "$.source");

        var requested = validation.Program!;
        var source = requested.Source!;
        var sourceSha256 = Sha256(sourceBytes);
        if (!source.Kind.Equals("pptx", StringComparison.Ordinal) ||
            !sourceSha256.Equals(source.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.source.hashMismatch",
                "The source-bound PPJ does not match the exact PPTX input bytes.",
                "$.source.sha256");

        var projected = PpjPresentationProjector.Project(
            sourceBytes,
            new PresentationProgramRequest
            {
                SourceUri = source.Uri,
                AssetRootUri = AssetRoot(requested),
                IncludeNodeMap = true,
            },
            limits);
        var baselineValidation = PpjProgramValidator.Validate(projected.Program.ProgramJson.Span);
        if (!baselineValidation.IsValid)
            throw new CodecException(
                "ppj.projection.invalid",
                "The exact PPTX source could not be reprojected into a valid PPJ baseline.",
                "$.source");
        var baseline = baselineValidation.Program!;

        RequireExactSourceDescriptor(baseline, requested);
        RequireBaselineAssets(baseline, requested);
        RequireRootTopology(baseline, requested);

        var artifact = projected.SourceArtifact;
        var presentation = artifact.Presentation ??
            throw new CodecException("ppj.source.presentation", "The exact source did not import as a Presentation artifact.", "$.source");
        var assetIds = BuildAssetCatalog(baseline, requested, projected, request.Assets, artifact);
        var changedNodeIds = new HashSet<string>(StringComparer.Ordinal);
        var physicalChanges = ApplyPages(baseline, requested, presentation, assetIds, changedNodeIds);

        byte[] output;
        IReadOnlyList<Diagnostic> diagnostics;
        if (!physicalChanges)
        {
            // A physical no-op is the original byte array, not a reserialized
            // equivalent package. PPJ-only intent/design metadata may still
            // advance outside the third-party PPTX.
            output = sourceBytes;
            diagnostics = projected.Diagnostics;
        }
        else
        {
            var exported = PptxCodec.Export(artifact, limits);
            output = exported.File;
            diagnostics = projected.Diagnostics.Concat(exported.Diagnostics).ToArray();
        }

        var receipt = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? projected.Program.NodeMapJson : ByteString.Empty,
            SourceSha256 = sourceSha256,
            OutputSha256 = Sha256(output),
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        receipt.Assets.Add(projected.Program.Assets.Select(asset => asset.Clone()));
        receipt.ChangedNodeIds.Add(changedNodeIds.OrderBy(id => id, StringComparer.Ordinal));
        receipt.ChangedParts.Add(ChangedParts(sourceBytes, output));
        return new(output, receipt, diagnostics);
    }

    private static void RequireExactSourceDescriptor(PpjProgramModel baseline, PpjProgramModel requested)
    {
        if (baseline.Source is null || requested.Source is null ||
            !JsonEqual(baseline.Root.GetProperty("source"), requested.Root.GetProperty("source")))
            throw new CodecException(
                "ppj.source.staleProjection",
                "The PPJ source descriptor no longer matches a fresh projection of the exact PPTX.",
                "$.source");
    }

    private static void RequireBaselineAssets(PpjProgramModel baseline, PpjProgramModel requested)
    {
        var requestedById = requested.Assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        foreach (var source in baseline.Assets)
        {
            if (!requestedById.TryGetValue(source.Id, out var current) ||
                !source.Uri.Equals(current.Uri, StringComparison.Ordinal) ||
                !source.MimeType.Equals(current.MimeType, StringComparison.OrdinalIgnoreCase) ||
                !source.Sha256.Equals(current.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "ppj.asset.sourceBindingChanged",
                    $"Source-projected PPJ asset {source.Id} changed its URI, MIME type, or content hash.",
                    "$.assets");
        }
    }

    private static void RequireRootTopology(PpjProgramModel baseline, PpjProgramModel requested)
    {
        if (requested.Components.Count != 0)
            throw Unsupported("$", "source-bound PPJ cannot introduce component definitions in this compiler slice");
        RequirePropertyEqual(baseline.Root, requested.Root, "sections", "$.sections");
        RequirePropertyEqual(baseline.Root, requested.Root, "customShows", "$.customShows");
        RequirePropertyEqual(baseline.Root, requested.Root, "comments", "$.comments");
    }

    private static Dictionary<string, string> BuildAssetCatalog(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PpjProjectionResult projected,
        IEnumerable<Asset> supplied,
        ArtifactEnvelope artifact)
    {
        var declared = requested.Assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        var suppliedById = supplied.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        foreach (var id in suppliedById.Keys.Where(id => !declared.ContainsKey(id)))
            throw new CodecException("ppj.asset.undeclared", $"PPJ compile received undeclared asset {id}.", "$.assets");

        var output = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var declaration in baseline.Assets)
        {
            var native = artifact.Assets.FirstOrDefault(asset =>
                AssetHash(asset).Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) &&
                asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase));
            if (native is null)
                throw new CodecException("ppj.asset.sourceMissing", $"Source-projected asset {declaration.Id} no longer resolves in the exact PPTX.", "$.assets");
            output[declaration.Id] = native.Id;
        }

        var baselineIds = baseline.Assets.Select(asset => asset.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var declaration in requested.Assets.Where(asset => !baselineIds.Contains(asset.Id)))
        {
            if (!suppliedById.TryGetValue(declaration.Id, out var asset) || asset.Data.IsEmpty)
                throw new CodecException("ppj.asset.missing", $"PPJ asset {declaration.Id} has no supplied bytes.", "$.assets");
            var hash = Sha256(asset.Data.Span);
            if (!hash.Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) ||
                !hash.Equals(asset.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.hashMismatch", $"PPJ asset {declaration.Id} does not match its declared SHA-256.", "$.assets");
            if (!asset.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase))
                throw new CodecException("ppj.asset.mimeMismatch", $"PPJ asset {declaration.Id} does not match its declared MIME type.", "$.assets");
            var nativeId = $"asset/ppj/{hash}";
            if (!artifact.Assets.Any(item => item.Id.Equals(nativeId, StringComparison.Ordinal)))
            {
                var copy = asset.Clone();
                copy.Id = nativeId;
                copy.FileName = declaration.Uri;
                copy.Sha256 = hash;
                artifact.Assets.Add(copy);
            }
            output[declaration.Id] = nativeId;
        }
        return output;
    }

    private static bool ApplyPages(
        PpjProgramModel baseline,
        PpjProgramModel requested,
        PresentationArtifact presentation,
        IReadOnlyDictionary<string, string> assets,
        ISet<string> changedNodeIds)
    {
        if (baseline.Pages.Count != requested.Pages.Count || baseline.Pages.Count != presentation.Slides.Count)
            throw Unsupported("$.pages", "source-bound page insertion or deletion requires an explicit page capability");
        var changed = false;
        for (var index = 0; index < baseline.Pages.Count; index++)
        {
            var before = baseline.Pages[index];
            var after = requested.Pages[index];
            var slide = presentation.Slides[index];
            var path = $"$.pages[{index}]";
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal))
                throw Unsupported(path, "source-bound page reorder or identity change");
            RequireNativeRef(before.Raw, after.Raw, path);
            RequireEqualExcept(before.Raw, after.Raw, path, "role", "claim", "elements");
            if (before.Elements.Count != after.Elements.Count || before.Elements.Count != slide.Elements.Count)
                throw Unsupported(path + ".elements", "source-bound element insertion or deletion requires an explicit capability");
            for (var elementIndex = 0; elementIndex < before.Elements.Count; elementIndex++)
            {
                var elementPath = $"{path}.elements[{elementIndex}]";
                if (ApplyElement(before.Elements[elementIndex], after.Elements[elementIndex], slide.Elements[elementIndex], assets, changedNodeIds, elementPath))
                {
                    changed = true;
                    changedNodeIds.Add(after.Id);
                }
            }
        }
        return changed;
    }

    private static bool ApplyElement(
        PpjElementModel before,
        PpjElementModel after,
        PresentationElement target,
        IReadOnlyDictionary<string, string> assets,
        ISet<string> changedNodeIds,
        string path)
    {
        if (!before.Id.Equals(after.Id, StringComparison.Ordinal) || !before.Type.Equals(after.Type, StringComparison.Ordinal))
            throw Unsupported(path, "source-bound element reorder, identity, or type change");
        RequireNativeRef(before.Raw, after.Raw, path);

        var changed = before switch
        {
            PpjTextElementModel beforeText when after is PpjTextElementModel afterText && target.ContentCase == PresentationElement.ContentOneofCase.Shape =>
                ApplyTextElement(beforeText, afterText, target.Shape, path),
            PpjShapeElementModel beforeShape when after is PpjShapeElementModel afterShape && target.ContentCase == PresentationElement.ContentOneofCase.Shape =>
                ApplyShapeElement(beforeShape, afterShape, target.Shape, path),
            PpjPlaceholderElementModel beforePlaceholder when after is PpjPlaceholderElementModel afterPlaceholder && target.ContentCase == PresentationElement.ContentOneofCase.Shape =>
                ApplyPlaceholderElement(beforePlaceholder, afterPlaceholder, target.Shape, path),
            PpjImageElementModel beforeImage when after is PpjImageElementModel afterImage && target.ContentCase == PresentationElement.ContentOneofCase.Image =>
                ApplyImageElement(beforeImage, afterImage, target.Image, assets, path),
            PpjChartElementModel beforeChart when after is PpjChartElementModel afterChart && target.ContentCase == PresentationElement.ContentOneofCase.Chart =>
                ApplyChartElement(beforeChart, afterChart, target.Chart, path),
            PpjTableElementModel beforeTable when after is PpjTableElementModel afterTable && target.ContentCase == PresentationElement.ContentOneofCase.Table =>
                ApplyTableElement(beforeTable, afterTable, target.Table, path),
            PpjConnectorElementModel beforeConnector when after is PpjConnectorElementModel afterConnector && target.ContentCase == PresentationElement.ContentOneofCase.Connector =>
                ApplyConnectorElement(beforeConnector, afterConnector, target.Connector, path),
            PpjGroupElementModel beforeGroup when after is PpjGroupElementModel afterGroup && target.ContentCase == PresentationElement.ContentOneofCase.Group =>
                ApplyGroupElement(beforeGroup, afterGroup, target.Group, assets, changedNodeIds, path),
            PpjOpaqueElementModel beforeOpaque when after is PpjOpaqueElementModel afterOpaque =>
                ApplyOpaqueElement(beforeOpaque, afterOpaque, target, path),
            _ => throw Unsupported(path, "the exact source object no longer matches its PPJ projection type"),
        };
        if (changed) changedNodeIds.Add(after.Id);
        return changed;
    }

    private static bool ApplyTextElement(PpjTextElementModel before, PpjTextElementModel after, PresentationShape target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "text", "fill", "stroke");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            changed |= ApplyText(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"), target, path + ".text");
        }
        changed |= ApplyFillProperty(before, after, target, "fill", path);
        changed |= ApplyStrokeProperty(before, after, target, "stroke", path);
        return changed;
    }

    private static bool ApplyShapeElement(PpjShapeElementModel before, PpjShapeElementModel after, PresentationShape target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "text", "style");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source text body");
            changed |= ApplyText(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"), target, path + ".text");
        }
        changed |= ApplyShapeStyle(before, after, target, path);
        return changed;
    }

    private static bool ApplyPlaceholderElement(PpjPlaceholderElementModel before, PpjPlaceholderElementModel after, PresentationShape target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "text");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "text"))
        {
            RequireCapability(after, "replaceText", path + ".text");
            if (before.Text is null || after.Text is null)
                throw Unsupported(path + ".text", "adding or removing a source placeholder text body");
            changed |= ApplyText(before.Text, after.Text, before.Raw.GetProperty("text"), after.Raw.GetProperty("text"), target, path + ".text");
        }
        return changed;
    }

    private static bool ApplyImageElement(
        PpjImageElementModel before,
        PpjImageElementModel after,
        PresentationImage target,
        IReadOnlyDictionary<string, string> assets,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "asset", "crop", "opacity");
        var changed = ApplyFrame(before, after, target, path);
        if (!before.AssetId.Equals(after.AssetId, StringComparison.Ordinal))
        {
            RequireCapability(after, "replaceImage", path + ".asset");
            if (!assets.TryGetValue(after.AssetId, out var nativeAssetId))
                throw new CodecException("ppj.asset.missing", $"PPJ image asset {after.AssetId} has no validated bytes.", path + ".asset");
            target.AssetId = nativeAssetId;
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "crop"))
        {
            RequireCapability(after, "setImageCrop", path + ".crop");
            target.Crop = after.Raw.TryGetProperty("crop", out var crop)
                ? new PresentationImageCrop
                {
                    LeftThousandthPercent = Crop(crop, "left"),
                    TopThousandthPercent = Crop(crop, "top"),
                    RightThousandthPercent = Crop(crop, "right"),
                    BottomThousandthPercent = Crop(crop, "bottom"),
                }
                : null;
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "opacity"))
        {
            RequireCapability(after, "setOpacity", path + ".opacity");
            if (after.Raw.TryGetProperty("opacity", out var opacity)) target.OpacityThousandthPercent = Unit(opacity.GetDouble());
            else target.ClearOpacityThousandthPercent();
            changed = true;
        }
        return changed;
    }

    private static bool ApplyChartElement(PpjChartElementModel before, PpjChartElementModel after, PresentationChart target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "title", "data");
        var changed = ApplyFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "title"))
        {
            RequireCapability(after, "setChartTitle", path + ".title");
            if (after.Title is not { PlainText: not null })
                throw Unsupported(path + ".title", "rich source chart-title authoring");
            target.Title = after.Title.PlainText;
            changed = true;
        }
        if (PropertyChanged(before.Raw, after.Raw, "data"))
        {
            RequireCapability(after, "setChartData", path + ".data");
            ApplyChartData(before, after, target, path + ".data");
            changed = true;
        }
        return changed;
    }

    private static bool ApplyTableElement(PpjTableElementModel before, PpjTableElementModel after, PresentationTable target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "rows");
        var changed = ApplyFrame(before, after, target, path);
        if (!PropertyChanged(before.Raw, after.Raw, "rows")) return changed;
        RequireCapability(after, "replaceText", path + ".rows");
        if (before.Rows.Count != after.Rows.Count || before.Rows.Count != target.Rows.Count)
            throw Unsupported(path + ".rows", "table row topology change");
        for (var row = 0; row < before.Rows.Count; row++)
        {
            var oldRow = before.Rows[row];
            var newRow = after.Rows[row];
            if (oldRow.Id != newRow.Id || oldRow.Height != newRow.Height || oldRow.Cells.Count != newRow.Cells.Count)
                throw Unsupported($"{path}.rows[{row}]", "table row topology or geometry change");
            for (var cell = 0; cell < oldRow.Cells.Count; cell++)
            {
                var oldCell = oldRow.Cells[cell];
                var newCell = newRow.Cells[cell];
                if (oldCell.Id != newCell.Id || oldCell.RowSpan != newCell.RowSpan || oldCell.ColumnSpan != newCell.ColumnSpan)
                    throw Unsupported($"{path}.rows[{row}].cells[{cell}]", "table cell topology change");
                if (TextEqual(oldCell.Text, newCell.Text)) continue;
                if (newCell.Text.PlainText is null || !TryProjectedCellCoordinates(oldCell.Id, out var physicalRow, out var physicalColumn) ||
                    physicalRow != row || physicalColumn >= target.Rows[row].Cells.Count)
                    throw Unsupported($"{path}.rows[{row}].cells[{cell}].text", "rich or noncanonical imported table-cell text");
                target.Rows[row].Cells[physicalColumn].Text = newCell.Text.PlainText;
                changed = true;
            }
        }
        return changed;
    }

    private static bool ApplyConnectorElement(PpjConnectorElementModel before, PpjConnectorElementModel after, PresentationConnector target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "stroke");
        var oldFrame = before.Frame;
        var changed = ApplyConnectorFrame(before, after, target, path);
        if (PropertyChanged(before.Raw, after.Raw, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".stroke");
            ApplyConnectorStroke(after.Raw.GetProperty("stroke"), target, path + ".stroke");
            changed = true;
        }
        _ = oldFrame;
        return changed;
    }

    private static bool ApplyGroupElement(
        PpjGroupElementModel before,
        PpjGroupElementModel after,
        PresentationGroup target,
        IReadOnlyDictionary<string, string> assets,
        ISet<string> changedNodeIds,
        string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame", "elements");
        var changed = ApplyFrame(before, after, target, path);
        if (before.Elements.Count != after.Elements.Count || before.Elements.Count != target.Children.Count)
            throw Unsupported(path + ".elements", "source-bound group child insertion or deletion");
        for (var index = 0; index < before.Elements.Count; index++)
            changed |= ApplyElement(before.Elements[index], after.Elements[index], target.Children[index], assets, changedNodeIds, $"{path}.elements[{index}]");
        return changed;
    }

    private static bool ApplyOpaqueElement(PpjOpaqueElementModel before, PpjOpaqueElementModel after, PresentationElement target, string path)
    {
        RequireEqualExcept(before.Raw, after.Raw, path, "role", "tags", "frame");
        if (!FrameChanged(before, after)) return false;
        return target.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => ApplyFrame(before, after, target.Shape, path),
            PresentationElement.ContentOneofCase.Image => ApplyFrame(before, after, target.Image, path),
            PresentationElement.ContentOneofCase.Chart => ApplyFrame(before, after, target.Chart, path),
            PresentationElement.ContentOneofCase.Table => ApplyFrame(before, after, target.Table, path),
            PresentationElement.ContentOneofCase.Connector => ApplyConnectorFrame(before, after, target.Connector, path),
            PresentationElement.ContentOneofCase.Group => ApplyFrame(before, after, target.Group, path),
            PresentationElement.ContentOneofCase.Opaque => ApplyFrame(before, after, target.Opaque, path),
            _ => throw Unsupported(path + ".frame", "placement of an unrecognized source object"),
        };
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationShape target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationImage target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationChart target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationTable target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationGroup target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyFrame(PpjElementModel before, PpjElementModel after, PresentationOpaqueElement target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        target.LeftEmu = Emu(after.Frame.X);
        target.TopEmu = Emu(after.Frame.Y);
        target.WidthEmu = Emu(after.Frame.Width);
        target.HeightEmu = Emu(after.Frame.Height);
        return true;
    }

    private static bool ApplyConnectorFrame(PpjElementModel before, PpjElementModel after, PresentationConnector target, string path)
    {
        if (!FrameChanged(before, after)) return false;
        RequireFrameChange(before, after, path);
        var old = before.Frame;
        var next = after.Frame;
        target.StartXEmu = TransformCoordinate(target.StartXEmu, old.X, old.Width, next.X, next.Width);
        target.EndXEmu = TransformCoordinate(target.EndXEmu, old.X, old.Width, next.X, next.Width);
        target.StartYEmu = TransformCoordinate(target.StartYEmu, old.Y, old.Height, next.Y, next.Height);
        target.EndYEmu = TransformCoordinate(target.EndYEmu, old.Y, old.Height, next.Y, next.Height);
        return true;
    }

    private static void RequireFrameChange(PpjElementModel before, PpjElementModel after, string path)
    {
        RequireCapability(after, "setFrame", path + ".frame");
        var oldFrame = before.Raw.GetProperty("frame");
        var newFrame = after.Raw.GetProperty("frame");
        RequireEqualExcept(oldFrame, newFrame, path + ".frame", "x", "y", "width", "height");
    }

    private static bool ApplyText(
        PpjTextContentModel before,
        PpjTextContentModel after,
        JsonElement beforeRaw,
        JsonElement afterRaw,
        PresentationShape target,
        string path)
    {
        if (JsonEqual(beforeRaw, afterRaw)) return false;
        if (!JsonEqual(MaskTextValues(beforeRaw), MaskTextValues(afterRaw)))
            throw Unsupported(path, "rich-text topology or styling change");
        if (target.TextBody is null)
            throw Unsupported(path, "text edit without one imported bounded text body");

        if (before.PlainText is not null || after.PlainText is not null)
        {
            if (before.PlainText is null || after.PlainText is null ||
                target.TextBody.Paragraphs.Count != 1 || target.TextBody.Paragraphs[0].Runs.Count != 1 ||
                target.TextBody.Paragraphs[0].Runs[0].ContentCase != PresentationTextRun.ContentOneofCase.Text)
                throw Unsupported(path, "plain/rich text conversion or multi-leaf plain replacement");
            target.TextBody.Paragraphs[0].Runs[0].Text = after.PlainText;
        }
        else
        {
            if (before.Paragraphs.Count != after.Paragraphs.Count || before.Paragraphs.Count != target.TextBody.Paragraphs.Count)
                throw Unsupported(path, "paragraph topology change");
            for (var paragraph = 0; paragraph < before.Paragraphs.Count; paragraph++)
            {
                if (before.Paragraphs[paragraph].Runs.Count != after.Paragraphs[paragraph].Runs.Count ||
                    before.Paragraphs[paragraph].Runs.Count != target.TextBody.Paragraphs[paragraph].Runs.Count)
                    throw Unsupported(path, "run topology change");
                for (var run = 0; run < before.Paragraphs[paragraph].Runs.Count; run++)
                {
                    var targetRun = target.TextBody.Paragraphs[paragraph].Runs[run];
                    if (targetRun.ContentCase != PresentationTextRun.ContentOneofCase.Text)
                        throw Unsupported(path, "non-text imported run mutation");
                    targetRun.Text = after.Paragraphs[paragraph].Runs[run].Text;
                }
            }
        }
        target.Text = string.Join("\n", target.TextBody.Paragraphs.Select(paragraph =>
            string.Concat(paragraph.Runs.Select(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Text ? run.Text : string.Empty))));
        return true;
    }

    private static bool ApplyShapeStyle(PpjElementModel before, PpjElementModel after, PresentationShape target, string path)
    {
        var oldStyle = OptionalProperty(before.Raw, "style");
        var newStyle = OptionalProperty(after.Raw, "style");
        if (JsonEqual(oldStyle, newStyle)) return false;
        if (oldStyle is { } oldValue && newStyle is { } newValue)
            RequireEqualExcept(oldValue, newValue, path + ".style", "fill", "stroke");
        else
        {
            var present = oldStyle ?? newStyle!.Value;
            foreach (var property in present.EnumerateObject())
                if (property.Name is not ("fill" or "stroke"))
                    throw Unsupported(path + ".style", $"changing {property.Name}");
        }
        var changed = false;
        if (PropertyChanged(oldStyle, newStyle, "fill"))
        {
            RequireCapability(after, "setFill", path + ".style.fill");
            ApplyFill(newStyle is { } style && style.TryGetProperty("fill", out var fill) ? fill : (JsonElement?)null, target, path + ".style.fill");
            changed = true;
        }
        if (PropertyChanged(oldStyle, newStyle, "stroke"))
        {
            RequireCapability(after, "setStroke", path + ".style.stroke");
            ApplyStroke(newStyle is { } style && style.TryGetProperty("stroke", out var stroke) ? stroke : (JsonElement?)null, target, path + ".style.stroke");
            changed = true;
        }
        return changed;
    }

    private static bool ApplyFillProperty(PpjElementModel before, PpjElementModel after, PresentationShape target, string name, string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setFill", $"{path}.{name}");
        ApplyFill(after.Raw.TryGetProperty(name, out var fill) ? fill : (JsonElement?)null, target, $"{path}.{name}");
        return true;
    }

    private static bool ApplyStrokeProperty(PpjElementModel before, PpjElementModel after, PresentationShape target, string name, string path)
    {
        if (!PropertyChanged(before.Raw, after.Raw, name)) return false;
        RequireCapability(after, "setStroke", $"{path}.{name}");
        ApplyStroke(after.Raw.TryGetProperty(name, out var stroke) ? stroke : (JsonElement?)null, target, $"{path}.{name}");
        return true;
    }

    private static void ApplyFill(JsonElement? fill, PresentationShape target, string path)
    {
        if (fill is null || fill.Value.GetProperty("type").GetString() == "none")
        {
            target.FillRgb = string.Empty;
            target.ClearFillOpacityThousandthPercent();
            return;
        }
        if (fill.Value.GetProperty("type").GetString() != "solid")
            throw Unsupported(path, "non-solid source-bound fill");
        target.FillRgb = Rgb(fill.Value.GetProperty("color"), path + ".color");
        if (fill.Value.TryGetProperty("opacity", out var opacity)) target.FillOpacityThousandthPercent = Unit(opacity.GetDouble());
        else target.ClearFillOpacityThousandthPercent();
    }

    private static void ApplyStroke(JsonElement? stroke, PresentationShape target, string path)
    {
        if (stroke is null)
        {
            target.LineRgb = string.Empty;
            target.LineStyle = "none";
            target.LineWidthEmu = 0;
            return;
        }
        target.LineRgb = Rgb(stroke.Value.GetProperty("color"), path + ".color");
        target.LineWidthEmu = Emu(stroke.Value.GetProperty("width").GetDouble());
        target.LineStyle = NativeDash(OptionalString(stroke.Value, "dash"));
        target.LineCap = OptionalString(stroke.Value, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke.Value, "join") ?? string.Empty;
    }

    private static void ApplyConnectorStroke(JsonElement stroke, PresentationConnector target, string path)
    {
        target.LineRgb = Rgb(stroke.GetProperty("color"), path + ".color");
        target.LineWidthEmu = Emu(stroke.GetProperty("width").GetDouble());
        target.LineStyle = NativeDash(OptionalString(stroke, "dash"));
        target.LineCap = OptionalString(stroke, "cap") ?? string.Empty;
        target.LineJoin = OptionalString(stroke, "join") ?? string.Empty;
    }

    private static void ApplyChartData(PpjChartElementModel before, PpjChartElementModel after, PresentationChart target, string path)
    {
        if (before.Data.Categories.Count != after.Data.Categories.Count ||
            before.Data.Series.Count != after.Data.Series.Count)
            throw Unsupported(path, "chart series or category topology change");
        var categories = after.Data.Categories.Select((item, index) => item.ValueKind == JsonValueKind.String
            ? item.GetString()!
            : throw Unsupported($"{path}.categories[{index}]", "non-string imported category")).ToArray();
        target.Categories.Clear();
        target.Categories.Add(categories);

        if (target.Type == SpreadsheetChartType.Combo)
        {
            if (target.ComboSeries.Count != after.Data.Series.Count)
                throw Unsupported(path, "combo-chart series topology change");
            for (var index = 0; index < after.Data.Series.Count; index++)
                ApplyChartSeries(before.Data.Series[index], after.Data.Series[index], target.ComboSeries[index].Series, $"{path}.series[{index}]");
        }
        else
        {
            if (target.Series.Count != after.Data.Series.Count)
                throw Unsupported(path, "chart series topology change");
            for (var index = 0; index < after.Data.Series.Count; index++)
                ApplyChartSeries(before.Data.Series[index], after.Data.Series[index], target.Series[index], $"{path}.series[{index}]");
        }
    }

    private static void ApplyChartSeries(PpjChartSeriesModel before, PpjChartSeriesModel after, SpreadsheetChartSeriesArtifact target, string path)
    {
        if (before.Id != after.Id || before.ChartType != after.ChartType || before.Axis != after.Axis || before.Values.Count != after.Values.Count)
            throw Unsupported(path, "chart-series identity or topology change");
        target.Name = after.Name;
        target.Values.Clear();
        foreach (var value in after.Values)
        {
            if (value is null) throw Unsupported(path, "null source-bound chart point");
            target.Values.Add(value.Value);
        }
    }

    private static bool FrameChanged(PpjElementModel before, PpjElementModel after) =>
        !JsonEqual(before.Raw.GetProperty("frame"), after.Raw.GetProperty("frame"));

    private static void RequireNativeRef(JsonElement before, JsonElement after, string path)
    {
        if (!before.TryGetProperty("nativeRef", out var oldReference) ||
            !after.TryGetProperty("nativeRef", out var newReference) ||
            !JsonEqual(oldReference, newReference))
            throw new CodecException(
                "ppj.nativeRef.stale",
                "The source-bound PPJ nativeRef is missing, changed, or no longer matches the exact source projection.",
                path + ".nativeRef");
    }

    private static void RequireCapability(PpjElementModel element, string operation, string path)
    {
        var nativeRef = element.NativeRef ?? throw new CodecException("ppj.nativeRef.missing", "Source-bound edits require a nativeRef.", path);
        var capability = nativeRef.Capabilities.FirstOrDefault(item => item.Operation.Equals(operation, StringComparison.Ordinal));
        if (capability is null || !capability.ExpectedHash.Equals(nativeRef.ObjectHash, StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "ppj.nativeRef.capabilityMissing",
                $"The exact source object did not issue the {operation} capability required by this edit.",
                path);
    }

    private static void RequireEqualExcept(JsonElement before, JsonElement after, string path, params string[] allowed)
    {
        var allowedNames = allowed.ToHashSet(StringComparer.Ordinal);
        var names = before.EnumerateObject().Select(property => property.Name)
            .Concat(after.EnumerateObject().Select(property => property.Name))
            .Distinct(StringComparer.Ordinal);
        foreach (var name in names)
        {
            if (allowedNames.Contains(name)) continue;
            var oldPresent = before.TryGetProperty(name, out var oldValue);
            var newPresent = after.TryGetProperty(name, out var newValue);
            if (oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue))
                throw Unsupported(path + "." + name, $"changing source-owned {name}");
        }
    }

    private static void RequirePropertyEqual(JsonElement before, JsonElement after, string name, string path)
    {
        var oldPresent = before.TryGetProperty(name, out var oldValue);
        var newPresent = after.TryGetProperty(name, out var newValue);
        if (oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue))
            throw Unsupported(path, "changing source-owned presentation topology");
    }

    private static bool PropertyChanged(JsonElement before, JsonElement after, string name)
    {
        var oldPresent = before.TryGetProperty(name, out var oldValue);
        var newPresent = after.TryGetProperty(name, out var newValue);
        return oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue);
    }

    private static bool PropertyChanged(JsonElement? before, JsonElement? after, string name)
    {
        var oldValue = default(JsonElement);
        var newValue = default(JsonElement);
        var oldPresent = before is { } oldObject && oldObject.TryGetProperty(name, out oldValue);
        var newPresent = after is { } newObject && newObject.TryGetProperty(name, out newValue);
        return oldPresent != newPresent || oldPresent && !JsonEqual(oldValue, newValue);
    }

    private static JsonElement? OptionalProperty(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value : null;

    private static JsonElement MaskTextValues(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.String)
        {
            using var plain = JsonDocument.Parse("\"\"");
            return plain.RootElement.Clone();
        }
        var bytes = PpjCanonicalJson.Write(value);
        var node = System.Text.Json.Nodes.JsonNode.Parse(bytes)!.AsObject();
        if (node["paragraphs"] is System.Text.Json.Nodes.JsonArray paragraphs)
            foreach (var paragraph in paragraphs.OfType<System.Text.Json.Nodes.JsonObject>())
                if (paragraph["runs"] is System.Text.Json.Nodes.JsonArray runs)
                    foreach (var run in runs.OfType<System.Text.Json.Nodes.JsonObject>()) run["text"] = string.Empty;
        using var document = JsonDocument.Parse(node.ToJsonString());
        return document.RootElement.Clone();
    }

    private static bool TextEqual(PpjTextContentModel left, PpjTextContentModel right)
    {
        if (left.PlainText is not null || right.PlainText is not null) return left.PlainText == right.PlainText;
        if (left.Paragraphs.Count != right.Paragraphs.Count) return false;
        for (var paragraph = 0; paragraph < left.Paragraphs.Count; paragraph++)
        {
            if (left.Paragraphs[paragraph].Runs.Count != right.Paragraphs[paragraph].Runs.Count) return false;
            for (var run = 0; run < left.Paragraphs[paragraph].Runs.Count; run++)
                if (left.Paragraphs[paragraph].Runs[run].Text != right.Paragraphs[paragraph].Runs[run].Text) return false;
        }
        return true;
    }

    private static bool TryProjectedCellCoordinates(string? id, out int row, out int column)
    {
        row = column = -1;
        if (id is null || !id.StartsWith("cell-", StringComparison.Ordinal)) return false;
        var values = id[5..].Split('-');
        return values.Length == 2 && int.TryParse(values[0], out var oneBasedRow) && int.TryParse(values[1], out var oneBasedColumn) &&
               (row = oneBasedRow - 1) >= 0 && (column = oneBasedColumn - 1) >= 0;
    }

    private static string AssetRoot(PpjProgramModel program)
    {
        var uri = program.Assets.FirstOrDefault()?.Uri;
        if (string.IsNullOrEmpty(uri)) return "deck.assets/media";
        var separator = uri.LastIndexOf('/');
        return separator > 0 ? uri[..separator] : "deck.assets/media";
    }

    private static IReadOnlyList<string> ChangedParts(byte[] before, byte[] after)
    {
        if (before.AsSpan().SequenceEqual(after)) return [];
        var oldParts = ZipHashes(before);
        var newParts = ZipHashes(after);
        return oldParts.Keys.Concat(newParts.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => !oldParts.TryGetValue(path, out var oldHash) || !newParts.TryGetValue(path, out var newHash) || oldHash != newHash)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static Dictionary<string, string> ZipHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var output = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(entry => !string.IsNullOrEmpty(entry.Name)))
        {
            using var input = entry.Open();
            using var memory = new MemoryStream();
            input.CopyTo(memory);
            output[entry.FullName] = Sha256(memory.ToArray());
        }
        return output;
    }

    private static string AssetHash(Asset asset) =>
        asset.Sha256 is { Length: 64 } ? asset.Sha256.ToLowerInvariant() : Sha256(asset.Data.Span);

    private static bool JsonEqual(JsonElement left, JsonElement right) =>
        PpjCanonicalJson.Write(left).AsSpan().SequenceEqual(PpjCanonicalJson.Write(right));

    private static bool JsonEqual(JsonElement? left, JsonElement? right) =>
        left is null || right is null ? left is null && right is null : JsonEqual(left.Value, right.Value);

    private static long Emu(double points) => checked((long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero));
    private static uint Unit(double value) => checked((uint)Math.Round(Math.Clamp(value, 0, 1) * 100_000, MidpointRounding.AwayFromZero));
    private static int Crop(JsonElement owner, string name) => owner.TryGetProperty(name, out var value)
        ? checked((int)Math.Round(Math.Clamp(value.GetDouble(), 0, 1) * 100_000, MidpointRounding.AwayFromZero))
        : 0;

    private static long TransformCoordinate(long source, double oldStart, double oldSize, double newStart, double newSize)
    {
        var sourcePoints = source / EmuPerPoint;
        var ratio = oldSize == 0 ? 0 : (sourcePoints - oldStart) / oldSize;
        return Emu(newStart + ratio * newSize);
    }

    private static string Rgb(JsonElement color, string path)
    {
        if (color.ValueKind != JsonValueKind.String) throw Unsupported(path, "theme or computed source-bound color");
        var value = color.GetString()!.TrimStart('#');
        if (value.Length != 6 || !value.All(Uri.IsHexDigit)) throw Unsupported(path, "non-RGB source-bound color");
        return value.ToUpperInvariant();
    }

    private static string NativeDash(string? value) => value switch
    {
        null or "" or "solid" => "solid",
        "dash" => "dashed",
        "dot" => "dotted",
        "dash-dot" => "dash-dot",
        "long-dash" => "long-dash",
        _ => throw Unsupported("stroke.dash", $"unsupported dash style {value}"),
    };

    private static string? OptionalString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out var value) ? value.GetString() : null;

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static CodecException Unsupported(string path, string detail) =>
        new("ppj.source.unsupportedMutation", $"Source-bound PPJ cannot safely compile {detail}.", path);
}
