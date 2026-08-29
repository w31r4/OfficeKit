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
    ArtifactEnvelope SourceArtifact);

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

    private static readonly HashSet<string> PresetGeometries = new(StringComparer.Ordinal)
    {
        "rect", "roundRect", "ellipse", "triangle", "rightTriangle", "diamond",
        "parallelogram", "trapezoid", "pentagon", "hexagon", "heptagon", "octagon",
        "chevron", "homePlate", "pie", "arc", "donut", "blockArc", "heart",
        "lightningBolt", "sun", "moon", "cloud", "star4", "star5", "star6",
        "star8", "star10", "star12", "leftArrow", "rightArrow", "upArrow",
        "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "bentArrow",
        "uturnArrow", "circularArrow", "wedgeRoundRectCallout", "wedgeEllipseCallout",
        "bracePair", "bracketPair", "flowChartProcess", "flowChartDecision",
        "flowChartData", "flowChartTerminator", "flowChartDocument", "flowChartPreparation",
    };

    internal static PpjProjectionResult Project(
        byte[] sourceBytes,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits)
    {
        var imported = PptxCodec.Import(sourceBytes, limits);
        var envelope = imported.Artifact;
        var presentation = envelope.Presentation ??
            throw new CodecException("ppj.projection.presentation", "The imported package did not produce a Presentation artifact.", "$");
        var sourceSha256 = Sha256(sourceBytes);
        var revision = $"pptx-{sourceSha256[..16]}";
        var sourceUri = string.IsNullOrWhiteSpace(request.SourceUri)
            ? $"deck.assets/source/{sourceSha256}.pptx"
            : request.SourceUri;
        var assetRoot = string.IsNullOrWhiteSpace(request.AssetRootUri)
            ? "deck.assets/media"
            : request.AssetRootUri.TrimEnd('/');

        var context = new ProjectionContext(sourceSha256, revision, assetRoot, envelope.Assets);
        RegisterIds(presentation, context);

        var pages = new JsonArray();
        foreach (var slide in presentation.Slides)
            pages.Add(ProjectPage(slide, presentation, context));

        var assets = context.ProgramAssets;
        var sections = ProjectSections(presentation, context);
        var customShows = ProjectCustomShows(presentation, context);
        var comments = ProjectComments(presentation, context);
        var nodeMap = context.BuildNodeMap();
        var nodeMapBytes = CanonicalBytes(nodeMap);
        var projectionPayload = new JsonObject
        {
            ["canvas"] = FrameDimensions(presentation),
            ["assets"] = assets.DeepClone(),
            ["pages"] = pages.DeepClone(),
            ["sections"] = sections.DeepClone(),
            ["customShows"] = customShows.DeepClone(),
            ["comments"] = comments.DeepClone(),
        };
        var projectionSha256 = Sha256(CanonicalBytes(projectionPayload));

        var root = new JsonObject
        {
            ["schema"] = "office-kit/ppj/v1",
            ["meta"] = new JsonObject
            {
                ["id"] = StableDocumentId(presentation.Id, sourceSha256),
                ["title"] = string.IsNullOrWhiteSpace(presentation.Name) ? "Imported presentation" : presentation.Name,
                ["language"] = "und",
                ["version"] = 1,
                ["description"] = "Source-derived PPJ projection. Unmodeled native content remains in the hash-bound PPTX source package.",
            },
            ["intent"] = ImportedIntent(),
            ["design"] = ImportedDesign(presentation),
            ["assets"] = assets,
            ["source"] = new JsonObject
            {
                ["kind"] = "pptx",
                ["uri"] = sourceUri,
                ["sha256"] = sourceSha256,
                ["revision"] = revision,
                ["projection"] = new JsonObject
                {
                    ["version"] = 1,
                    ["sha256"] = projectionSha256,
                    ["nodeMapSha256"] = Sha256(nodeMapBytes),
                    ["visibleObjectCount"] = context.VisibleObjectCount,
                },
            },
            ["pages"] = pages,
        };
        if (sections.Count > 0) root["sections"] = sections;
        if (customShows.Count > 0) root["customShows"] = customShows;
        if (comments.Count > 0) root["comments"] = comments;

        var candidateBytes = CanonicalBytes(root);
        var validation = PpjProgramValidator.Validate(candidateBytes);
        if (!validation.IsValid)
        {
            var first = validation.Diagnostics[0];
            throw new CodecException(first.Code, first.Message, first.Path);
        }

        var result = new PresentationProgramResult
        {
            ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
            ProgramSha256 = validation.ProgramSha256,
            NodeMapJson = request.IncludeNodeMap ? ByteString.CopyFrom(nodeMapBytes) : ByteString.Empty,
            SourceSha256 = sourceSha256,
            SourceBound = true,
            ExpandedElementCount = checked((uint)validation.Expansion!.ExpandedElementCount),
        };
        result.Assets.Add(context.ResultAssets.Select(asset => asset.Clone()));
        return new(result, imported.Diagnostics, envelope);
    }

    private static JsonObject ImportedIntent() => new()
    {
        ["brief"] = new JsonObject
        {
            ["primaryJob"] = "handoff",
            ["expectedOutcome"] = "Continue editing this presentation while preserving source-owned native content.",
            ["evidenceBoundary"] = "The source presentation does not declare its audience, factual authority, or intended outcome.",
        },
        ["audience"] = new JsonObject
        {
            ["description"] = "Imported presentation audience was not declared.",
        },
        ["narrative"] = new JsonObject
        {
            ["thesis"] = "Preserve and continue the imported presentation without inventing its original intent.",
        },
        ["editorial"] = new JsonObject
        {
            ["tone"] = new JsonArray("source-derived"),
            ["avoid"] = new JsonArray("Inventing missing facts or silently rewriting source-owned content"),
        },
        ["delivery"] = new JsonObject
        {
            ["mode"] = "hybrid",
            ["mediumFit"] = "acceptable",
            ["mediumFitNote"] = "Delivery intent was not declared in the imported file.",
        },
    };

    private static JsonObject ImportedDesign(PresentationArtifact presentation) => new()
    {
        ["canvas"] = FrameDimensions(presentation),
        ["theme"] = new JsonObject
        {
            ["name"] = "Source-owned presentation theme",
            ["colors"] = new JsonArray(new JsonObject
            {
                ["id"] = "source-neutral",
                ["value"] = "#000000",
                ["role"] = "Fallback only; imported native styling remains source-owned",
            }),
        },
        ["fonts"] = new JsonArray(new JsonObject
        {
            ["id"] = "source-font",
            ["family"] = "Arial",
            ["language"] = "und",
        }),
        ["styles"] = new JsonObject(),
        ["grammar"] = new JsonObject
        {
            ["name"] = "Source-derived projection",
            ["rationale"] = "The original PPTX remains the authority for native visual styling that PPJ does not model.",
            ["visualThesis"] = "Preserve the imported design and expose only bounded semantic edits.",
            ["surfaceHierarchy"] = new JsonArray("Keep source-owned surfaces unchanged unless a capability explicitly permits an edit."),
            ["typographyRhythm"] = new JsonArray("Retain imported typography through the source package."),
            ["geometryRules"] = new JsonArray("Retain imported geometry and z-order unless a nativeRef capability permits a change."),
            ["densityRhythm"] = new JsonArray("Retain the source page density."),
            ["carrierRules"] = new JsonArray("Use projected typed objects where safe and opaque native objects everywhere else."),
            ["forbiddenPatterns"] = new JsonArray("Rebuilding opaque content", "Guessing unsupported native semantics"),
        },
        ["motionPolicy"] = "explicit",
    };

    private static JsonObject FrameDimensions(PresentationArtifact presentation) => new()
    {
        ["width"] = Points(presentation.SlideWidthEmu),
        ["height"] = Points(presentation.SlideHeightEmu),
        ["unit"] = "pt",
    };

    private static void RegisterIds(PresentationArtifact presentation, ProjectionContext context)
    {
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.RegisterPage(slide.Id);
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
        var pageHash = HashOrFallback(slide.Source?.SlideXmlSha256, slide.ToByteArray());
        var pageCapabilities = new List<CapabilitySpec>();
        if (slide.Source?.DeletionCapability?.Supported == true)
            pageCapabilities.Add(new("delete", ["element"]));
        // A native slide clone needs fresh page/element identities and a
        // complete source-owned subtree mapping. Do not issue the underlying
        // codec capability until PPJ can represent that bounded clone request.

        var elements = new JsonArray();
        foreach (var element in slide.Elements)
            elements.Add(ProjectElement(element, pageId, context));

        var page = new JsonObject
        {
            ["id"] = pageId,
            ["name"] = string.IsNullOrWhiteSpace(slide.Name) ? null : slide.Name,
            ["role"] = "source continuation",
            ["elements"] = elements,
            ["nativeRef"] = NativeRef(context, $"page:{pageId}", pageHash, pageCapabilities),
        };
        if (slide.HasHidden) page["hidden"] = slide.Hidden;
        if (!string.IsNullOrEmpty(slide.SpeakerNotes?.Text)) page["notes"] = slide.SpeakerNotes.Text;
        if (ProjectBackground(slide.Background, context) is { } background) page["background"] = background;

        var animations = ProjectAnimations(slide, context);
        if (animations.Count > 0) page["animations"] = animations;
        if (ProjectTransition(slide, presentation, context) is { } transition) page["transition"] = transition;
        return page;
    }

    private static JsonObject ProjectElement(PresentationElement element, string pageId, ProjectionContext context)
    {
        var id = context.ElementId(pageId, element.Id);
        var hash = HashOrFallback(element.Source?.ElementSha256, element.ToByteArray());
        var capabilities = Capabilities(element);
        var nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, capabilities);
        JsonObject projected = element.ContentCase switch
        {
            PresentationElement.ContentOneofCase.Shape => ProjectShape(element, id, nativeRef),
            PresentationElement.ContentOneofCase.Image => ProjectImage(element, id, nativeRef, context),
            PresentationElement.ContentOneofCase.Table => ProjectTable(element, id, nativeRef),
            PresentationElement.ContentOneofCase.Connector => ProjectConnector(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Chart => ProjectChart(element, id, nativeRef),
            PresentationElement.ContentOneofCase.Group => ProjectGroup(element, id, nativeRef, pageId, context),
            PresentationElement.ContentOneofCase.Opaque => ProjectOpaque(element, id, nativeRef),
            _ => ProjectOpaque(element, id, nativeRef, "unknown"),
        };
        if (projected["type"]!.GetValue<string>() == "opaque")
        {
            nativeRef = NativeRef(context, $"element:{pageId}:{id}", hash, OpaqueCapabilities(element));
            projected["nativeRef"] = nativeRef;
        }
        context.RecordNode(pageId, id, projected["type"]!.GetValue<string>(), nativeRef);
        return projected;
    }

    private static JsonObject ProjectShape(PresentationElement element, string id, JsonObject nativeRef)
    {
        var shape = element.Shape;
        var frame = ShapeFrame(shape);
        var text = TextContent(shape.TextBody, shape.Text);
        var hasText = !string.IsNullOrEmpty(shape.Text) || shape.TextBody?.Paragraphs.Count > 0;
        var isPlaceholder = shape.Placeholder is not null;
        var isTextBox = shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry);
        if (!isPlaceholder && !isTextBox && !PresetGeometries.Contains(shape.Geometry))
            return ProjectOpaque(element, id, nativeRef, "shape", $"Preserved source shape with unsupported geometry '{shape.Geometry}'.");

        var common = ElementBase(id, element.Name, frame, Accessibility(shape.Accessibility), nativeRef);
        if (shape.Placeholder is not null)
        {
            common["type"] = "placeholder";
            common["placeholderType"] = PlaceholderType(shape.Placeholder.Type);
            common["index"] = shape.Placeholder.Index;
            if (hasText) common["text"] = text;
            return common;
        }
        if (shape.Geometry is "textbox" or "none" || string.IsNullOrEmpty(shape.Geometry))
        {
            common["type"] = "text";
            common["text"] = text;
            ApplyTextContainerStyle(common, shape);
            return common;
        }
        common["type"] = "shape";
        common["geometry"] = new JsonObject { ["kind"] = "preset", ["preset"] = shape.Geometry };
        if (hasText) common["text"] = text;
        var style = ShapeStyle(shape);
        if (style.Count > 0) common["style"] = style;
        return common;
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
        var output = ElementBase(id, element.Name, ImageFrame(image), ImageAccessibility(image), nativeRef);
        output["type"] = "image";
        output["asset"] = assetId;
        output["fit"] = "stretch";
        if (image.Crop is not null)
        {
            output["crop"] = new JsonObject
            {
                ["left"] = Crop(image.Crop.LeftThousandthPercent),
                ["top"] = Crop(image.Crop.TopThousandthPercent),
                ["right"] = Crop(image.Crop.RightThousandthPercent),
                ["bottom"] = Crop(image.Crop.BottomThousandthPercent),
            };
        }
        if (image.HasOpacityThousandthPercent)
            output["opacity"] = Unit(image.OpacityThousandthPercent);
        if (!string.IsNullOrEmpty(image.MaskPreset) && PresetGeometries.Contains(image.MaskPreset))
            output["mask"] = new JsonObject { ["kind"] = "preset", ["preset"] = image.MaskPreset };
        if (image.Border is not null && !string.IsNullOrEmpty(image.Border.ColorRgb))
            output["border"] = Stroke(image.Border.ColorRgb, image.Border.WidthEmu, image.Border.Style, image.Border.Cap, image.Border.Join,
                image.Border.HasOpacityThousandthPercent ? Unit(image.Border.OpacityThousandthPercent) : null);
        if (image.Shadow is not null && !string.IsNullOrEmpty(image.Shadow.ColorRgb))
            output["shadow"] = Shadow(image.Shadow);
        return output;
    }

    private static JsonObject ProjectChart(PresentationElement element, string id, JsonObject nativeRef)
    {
        var chart = element.Chart;
        var type = ChartType(chart.Type);
        var series = chart.Type == SpreadsheetChartType.Combo
            ? chart.ComboSeries.Select(item => item.Series).ToArray()
            : chart.Series.ToArray();
        if (type is null || series.Length == 0 || series.Any(item => item.Values.Count != chart.Categories.Count))
            return ProjectOpaque(element, id, nativeRef, "chart", "Preserved source chart outside the bounded PPJ data profile.");

        var output = ElementBase(id, element.Name, ChartFrame(chart), Accessibility(chart.Accessibility), nativeRef);
        output["type"] = "chart";
        output["chartType"] = type;
        if (!string.IsNullOrEmpty(chart.Title)) output["title"] = chart.Title;
        var categories = new JsonArray();
        foreach (var value in chart.Categories) categories.Add(value);
        var seriesJson = new JsonArray();
        for (var index = 0; index < series.Length; index++)
        {
            var item = series[index];
            var values = new JsonArray();
            foreach (var value in item.Values) values.Add(value);
            var entry = new JsonObject
            {
                ["id"] = $"series-{index + 1}",
                ["name"] = item.Name ?? string.Empty,
                ["values"] = values,
            };
            if (chart.Type == SpreadsheetChartType.Combo)
            {
                entry["chartType"] = ChartType(chart.ComboSeries[index].Type) ?? "line";
                entry["axis"] = chart.ComboSeries[index].AxisGroup == PresentationChartAxisGroup.Secondary ? "secondary" : "primary";
            }
            seriesJson.Add(entry);
        }
        output["data"] = new JsonObject { ["categories"] = categories, ["series"] = seriesJson };
        return output;
    }

    private static JsonObject ProjectTable(PresentationElement element, string id, JsonObject nativeRef)
    {
        var table = element.Table;
        if (table.ColumnWidthsEmu.Count == 0 || table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != table.ColumnWidthsEmu.Count))
            return ProjectOpaque(element, id, nativeRef, "table", "Preserved source table outside the bounded rectangular grid profile.");

        var output = ElementBase(id, element.Name, TableFrame(table), Accessibility(table.Accessibility), nativeRef);
        output["type"] = "table";
        var columns = new JsonArray();
        for (var index = 0; index < table.ColumnWidthsEmu.Count; index++)
            columns.Add(new JsonObject { ["id"] = $"column-{index + 1}", ["width"] = Points(table.ColumnWidthsEmu[index]) });
        output["columns"] = columns;
        output["rows"] = ProjectTableRows(table);
        var style = new JsonObject();
        if (table.HasFirstRow) style["headerRows"] = table.FirstRow ? 1 : 0;
        if (table.HasBandedRows) style["bandedRows"] = table.BandedRows;
        if (table.HasBandedColumns) style["bandedColumns"] = table.BandedColumns;
        if (table.HasFirstColumn) style["firstColumnEmphasis"] = table.FirstColumn;
        if (table.HasLastColumn) style["lastColumnEmphasis"] = table.LastColumn;
        if (style.Count > 0) output["style"] = style;
        return output;
    }

    private static JsonArray ProjectTableRows(PresentationTable table)
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
                var cell = new JsonObject
                {
                    ["id"] = $"cell-{rowIndex + 1}-{columnIndex + 1}",
                    ["text"] = table.Rows[rowIndex].Cells[columnIndex].Text ?? string.Empty,
                };
                if (mergeByOrigin.TryGetValue((rowIndex, columnIndex), out var merge))
                {
                    cell["rowSpan"] = checked((int)(merge.EndRow - merge.StartRow + 1));
                    cell["columnSpan"] = checked((int)(merge.EndColumn - merge.StartColumn + 1));
                }
                cells.Add(cell);
            }
            rows.Add(new JsonObject
            {
                ["id"] = $"row-{rowIndex + 1}",
                ["height"] = Points(table.Rows[rowIndex].HeightEmu),
                ["cells"] = cells,
            });
        }
        return rows;
    }

    private static JsonObject ProjectConnector(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context)
    {
        var connector = element.Connector;
        var output = ElementBase(id, element.Name, ConnectorFrame(connector), Accessibility(connector.Accessibility), nativeRef);
        output["type"] = "connector";
        output["connectorType"] = connector.ConnectorType is "elbow" or "curved" ? connector.ConnectorType : "straight";
        output["from"] = ConnectorEndpoint(connector.StartTargetId, connector.StartXEmu, connector.StartYEmu, pageId, context);
        output["to"] = ConnectorEndpoint(connector.EndTargetId, connector.EndXEmu, connector.EndYEmu, pageId, context);
        output["stroke"] = Stroke(
            string.IsNullOrEmpty(connector.LineRgb) ? "000000" : connector.LineRgb,
            connector.LineWidthEmu,
            connector.LineStyle,
            connector.LineCap,
            connector.LineJoin,
            null);
        if (Arrow(connector.StartArrow) is { } startArrow) output["startArrow"] = startArrow;
        if (Arrow(connector.EndArrow) is { } endArrow) output["endArrow"] = endArrow;
        return output;
    }

    private static JsonObject ProjectGroup(
        PresentationElement element,
        string id,
        JsonObject nativeRef,
        string pageId,
        ProjectionContext context)
    {
        var group = element.Group;
        if (group.Children.Count == 0)
            return ProjectOpaque(element, id, nativeRef, "group", "Preserved empty or unsupported native group.");
        var output = ElementBase(id, element.Name, GroupFrame(group), Accessibility(group.Accessibility), nativeRef);
        output["type"] = "group";
        var children = new JsonArray();
        foreach (var child in group.Children) children.Add(ProjectElement(child, pageId, context));
        output["elements"] = children;
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
        var output = ElementBase(id, element.Name, ElementFrame(element), null, nativeRef);
        output["type"] = "opaque";
        output["nativeKind"] = nativeKind;
        output["summary"] = summary ?? $"Preserved source-owned {nativeKind} object; only issued nativeRef capabilities are editable.";
        if (opaque is not null && PpjNativeTextProjection.TryRead(opaque.RawXml, out var nativeLeaves))
        {
            var visible = new JsonArray();
            foreach (var leaf in nativeLeaves) visible.Add(leaf);
            output["visibleText"] = visible;
        }
        else if (!string.IsNullOrWhiteSpace(opaque?.Text))
        {
            var visible = new JsonArray();
            foreach (var line in opaque.Text.Split('\n').Where(line => line.Length > 0)) visible.Add(line);
            if (visible.Count > 0) output["visibleText"] = visible;
        }
        return output;
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
            ["id"] = id,
            ["frame"] = frame,
            ["nativeRef"] = nativeRef,
        };
        if (!string.IsNullOrWhiteSpace(name)) output["name"] = name;
        if (accessibility is not null) output["accessibility"] = accessibility;
        return output;
    }

    private static JsonObject ShapeStyle(PresentationShape shape)
    {
        var style = new JsonObject();
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = "solid", ["color"] = Color(shape.FillRgb) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = Unit(shape.FillOpacityThousandthPercent);
            style["fill"] = fill;
        }
        if (!string.IsNullOrEmpty(shape.LineRgb) && shape.LineStyle != "none")
            style["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin, null);
        if (shape.Shadow is not null && !string.IsNullOrEmpty(shape.Shadow.ColorRgb))
            style["shadow"] = Shadow(shape.Shadow);
        return style;
    }

    private static JsonNode TextContent(PresentationTextBody? body, string? fallback)
    {
        if (body is null || body.Paragraphs.Count == 0 ||
            body.Paragraphs.Any(paragraph => paragraph.Runs.Count == 0 ||
                paragraph.Runs.Any(run => run.ContentCase != PresentationTextRun.ContentOneofCase.Text)))
            return JsonValue.Create(fallback ?? string.Empty)!;

        var paragraphs = new JsonArray();
        for (var paragraphIndex = 0; paragraphIndex < body.Paragraphs.Count; paragraphIndex++)
        {
            var source = body.Paragraphs[paragraphIndex];
            var paragraph = new JsonObject
            {
                ["id"] = $"paragraph-{paragraphIndex + 1}",
            };
            var paragraphStyle = new JsonObject();
            if (source.HasLevel) paragraphStyle["level"] = checked((int)source.Level);
            if (source.HasAlignment && ParagraphAlignment(source.Alignment) is { } alignment)
                paragraphStyle["alignment"] = alignment;
            if (source.LeftMarginCase == PresentationTextParagraph.LeftMarginOneofCase.MarginLeftEmu)
                paragraphStyle["indent"] = Points(source.MarginLeftEmu);
            if (source.IndentationCase == PresentationTextParagraph.IndentationOneofCase.IndentEmu)
                paragraphStyle["hanging"] = -Points(source.IndentEmu);
            if (source.LineSpacingCase == PresentationTextParagraph.LineSpacingOneofCase.LineSpacingPoints)
                paragraphStyle["lineSpacing"] = Math.Max(0.001, source.LineSpacingPoints);
            if (source.SpaceBeforeCase == PresentationTextParagraph.SpaceBeforeOneofCase.SpaceBeforePoints)
                paragraphStyle["spaceBefore"] = Math.Max(0, source.SpaceBeforePoints);
            if (source.SpaceAfterCase == PresentationTextParagraph.SpaceAfterOneofCase.SpaceAfterPoints)
                paragraphStyle["spaceAfter"] = Math.Max(0, source.SpaceAfterPoints);
            if (paragraphStyle.Count > 0) paragraph["style"] = paragraphStyle;

            var runs = new JsonArray();
            for (var runIndex = 0; runIndex < source.Runs.Count; runIndex++)
            {
                var sourceRun = source.Runs[runIndex];
                var run = new JsonObject
                {
                    ["id"] = $"run-{paragraphIndex + 1}-{runIndex + 1}",
                    ["text"] = sourceRun.Text,
                };
                var style = RunStyle(sourceRun);
                if (style.Count > 0) run["style"] = style;
                runs.Add(run);
            }
            paragraph["runs"] = runs;
            paragraphs.Add(paragraph);
        }
        return new JsonObject { ["paragraphs"] = paragraphs };
    }

    private static JsonObject RunStyle(PresentationTextRun run)
    {
        var style = new JsonObject();
        if (run.HasFontFamily) style["fontFamily"] = run.FontFamily;
        if (run.HasFontSizePoints && run.FontSizePoints > 0) style["size"] = run.FontSizePoints;
        if (run.HasBold) style["bold"] = run.Bold;
        if (run.HasItalic) style["italic"] = run.Italic;
        if (run.HasColorRgb && !string.IsNullOrEmpty(run.ColorRgb)) style["color"] = Color(run.ColorRgb);
        return style;
    }

    private static string? ParagraphAlignment(string value) => value switch
    {
        "l" or "left" => "left",
        "ctr" or "center" => "center",
        "r" or "right" => "right",
        "just" or "justify" => "justify",
        "dist" or "distributed" => "distributed",
        _ => null,
    };

    private static void ApplyTextContainerStyle(JsonObject output, PresentationShape shape)
    {
        if (!string.IsNullOrEmpty(shape.FillRgb))
        {
            var fill = new JsonObject { ["type"] = "solid", ["color"] = Color(shape.FillRgb) };
            if (shape.HasFillOpacityThousandthPercent) fill["opacity"] = Unit(shape.FillOpacityThousandthPercent);
            output["fill"] = fill;
        }
        if (!string.IsNullOrEmpty(shape.LineRgb) && shape.LineStyle != "none")
            output["stroke"] = Stroke(shape.LineRgb, shape.LineWidthEmu, shape.LineStyle, shape.LineCap, shape.LineJoin, null);
    }

    private static JsonObject? ProjectBackground(PresentationBackground? background, ProjectionContext context)
    {
        if (background is null) return null;
        if (!string.IsNullOrEmpty(background.ImageAssetId) && context.TryMaterializeAsset(background.ImageAssetId, out var assetId))
            return new JsonObject { ["type"] = "image", ["asset"] = assetId, ["fit"] = "stretch" };
        if (!string.IsNullOrEmpty(background.ColorRgb))
            return new JsonObject { ["type"] = "solid", ["color"] = Color(background.ColorRgb) };
        return null;
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
                ["id"] = context.UniqueId($"animation-{animation.Id}"),
                ["target"] = targetId,
                ["phase"] = animation.Phase,
                ["effect"] = animation.Effect,
                ["start"] = animation.Start,
                ["durationMs"] = animation.HasDurationMs ? checked((int)animation.DurationMs) : 500,
            };
            if (!string.IsNullOrEmpty(animation.Direction)) item["direction"] = animation.Direction;
            if (animation.HasDelayMs) item["delayMs"] = checked((int)animation.DelayMs);
            if (!string.IsNullOrEmpty(animation.TextBuild)) item["textBuild"] = animation.TextBuild;
            if (!string.IsNullOrEmpty(animation.ChartBuild)) item["chartBuild"] = animation.ChartBuild;
            if (animation.HasStaggerMs) item["staggerMs"] = checked((int)animation.StaggerMs);
            if (animation.HasAnimateChartBackground) item["animateChartBackground"] = animation.AnimateChartBackground;
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
                pairs.Add(new JsonObject { ["key"] = context.UniqueId($"morph-{pair.Key}"), ["from"] = fromId, ["to"] = toId });
            }
            if (pairs.Count == 0) return null;
            return new JsonObject
            {
                ["type"] = "morph",
                ["durationMs"] = slide.Morph.HasDurationMs ? checked((int)slide.Morph.DurationMs) : 800,
                ["fromPage"] = context.PageId(fromPage.Id),
                ["morphPairs"] = pairs,
            };
        }
        var transition = slide.Transition;
        if (transition is null || transition.Effect is not ("fade" or "push" or "wipe")) return null;
        var output = new JsonObject { ["type"] = transition.Effect };
        if (transition.HasDurationMs) output["durationMs"] = checked((int)transition.DurationMs);
        if (transition.Effect is "push" or "wipe" && transition.Direction is "left" or "right" or "up" or "down")
            output["direction"] = transition.Direction;
        return output;
    }

    private static JsonArray ProjectSections(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var section in presentation.Sections)
        {
            var pages = new JsonArray();
            foreach (var id in section.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(pageId);
            if (pages.Count == 0) continue;
            output.Add(new JsonObject
            {
                ["id"] = context.UniqueId($"section-{section.Id}"),
                ["name"] = string.IsNullOrWhiteSpace(section.Name) ? "Section" : section.Name,
                ["pages"] = pages,
            });
        }
        return output;
    }

    private static JsonArray ProjectCustomShows(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var show in presentation.CustomShows)
        {
            var pages = new JsonArray();
            foreach (var id in show.SlideIds)
                if (context.TryPageId(id, out var pageId)) pages.Add(pageId);
            if (pages.Count == 0) continue;
            output.Add(new JsonObject
            {
                ["id"] = context.UniqueId($"show-{show.Id}"),
                ["name"] = string.IsNullOrWhiteSpace(show.Name) ? "Custom show" : show.Name,
                ["pages"] = pages,
            });
        }
        return output;
    }

    private static JsonArray ProjectComments(PresentationArtifact presentation, ProjectionContext context)
    {
        var output = new JsonArray();
        foreach (var slide in presentation.Slides)
        {
            var pageId = context.PageId(slide.Id);
            foreach (var comment in slide.LegacyComments)
            {
                if (!DateTimeOffset.TryParse(comment.CreatedAt, out var createdAt) || string.IsNullOrWhiteSpace(comment.Text)) continue;
                output.Add(new JsonObject
                {
                    ["id"] = context.UniqueId($"comment-{pageId}-{comment.Id}"),
                    ["page"] = pageId,
                    ["author"] = string.IsNullOrWhiteSpace(comment.Author) ? "Unknown author" : comment.Author,
                    ["text"] = comment.Text,
                    ["createdAt"] = createdAt.ToUniversalTime().ToString("O"),
                    ["resolved"] = false,
                    ["position"] = new JsonObject { ["x"] = Points(comment.PositionXEmu), ["y"] = Points(comment.PositionYEmu) },
                });
            }
        }
        return output;
    }

    private static IReadOnlyList<CapabilitySpec> Capabilities(PresentationElement element)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        switch (element.ContentCase)
        {
            case PresentationElement.ContentOneofCase.Shape:
                if (source.TextEditable && TextTopologyRepresentable(element.Shape.TextBody))
                    output.Add(new("replaceText", ["text"]));
                if (source.Editable)
                {
                    output.Add(new("setFill", ["fill"]));
                    output.Add(new("setStroke", ["stroke"]));
                    output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                }
                break;
            case PresentationElement.ContentOneofCase.Image when source.Editable:
                output.Add(new("replaceImage", ["image.asset"]));
                output.Add(new("setImageCrop", ["image.crop"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                output.Add(new("setOpacity", ["opacity"]));
                break;
            case PresentationElement.ContentOneofCase.Chart when source.Editable:
                output.Add(new("setChartTitle", ["chart.title"]));
                output.Add(new("setChartData", ["chart.data"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Table when source.Editable:
                output.Add(new("replaceText", ["text"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Connector when source.Editable:
                output.Add(new("setStroke", ["stroke"]));
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Group when source.Editable:
                output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
            case PresentationElement.ContentOneofCase.Opaque:
                if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
                break;
        }
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static bool TextTopologyRepresentable(PresentationTextBody? body) =>
        body is not null && body.Paragraphs.Count > 0 &&
        body.Paragraphs.All(paragraph => paragraph.Runs.Count > 0 &&
            paragraph.Runs.All(run => run.ContentCase == PresentationTextRun.ContentOneofCase.Text));

    private static IReadOnlyList<CapabilitySpec> OpaqueCapabilities(PresentationElement element)
    {
        var output = new List<CapabilitySpec>();
        var source = element.Source;
        if (source is null) return output;
        if (element.ContentCase == PresentationElement.ContentOneofCase.Opaque &&
            PpjNativeTextProjection.TryRead(element.Opaque.RawXml, out _))
            output.Add(new("replaceText", ["visibleText"]));
        if (source.Editable) output.Add(new("setFrame", ["frame.x", "frame.y", "frame.width", "frame.height"]));
        // Diagram, OLE, and source-owned chart payloads stay opaque until the
        // corresponding PPJ typed state is projected. A runtime can edit such
        // objects only after the public language has somewhere to represent
        // both the old and requested value.
        if (source.DeletionCapability?.Supported == true) output.Add(new("delete", ["element"]));
        if (source.ZOrderCapability?.Supported == true) output.Add(new("reorder", ["zOrder"]));
        return output;
    }

    private static JsonObject NativeRef(
        ProjectionContext context,
        string scope,
        string objectHash,
        IEnumerable<CapabilitySpec> capabilities)
    {
        var capabilityArray = new JsonArray();
        foreach (var capability in capabilities
                     .OrderBy(item => item.Operation, StringComparer.Ordinal)
                     .ThenBy(item => string.Join("\0", item.Fields), StringComparer.Ordinal))
        {
            var fields = new JsonArray();
            foreach (var field in capability.Fields) fields.Add(field);
            capabilityArray.Add(new JsonObject
            {
                ["id"] = $"cap-{capability.Operation}-{Sha256(Encoding.UTF8.GetBytes(scope + capability.Operation))[..10]}",
                ["operation"] = capability.Operation,
                ["expectedHash"] = objectHash,
                ["fields"] = fields,
            });
        }
        return new JsonObject
        {
            ["handle"] = $"nr-{Sha256(Encoding.UTF8.GetBytes(context.SourceSha256 + "\0" + scope))}",
            ["sourceSha256"] = context.SourceSha256,
            ["revision"] = context.Revision,
            ["objectHash"] = objectHash,
            ["capabilitySetSha256"] = Sha256(CanonicalBytes(capabilityArray)),
            ["capabilities"] = capabilityArray,
        };
    }

    private static JsonObject ShapeFrame(PresentationShape shape)
    {
        var frame = Frame(shape.LeftEmu, shape.TopEmu, shape.WidthEmu, shape.HeightEmu);
        if (shape.Transform?.HasRotationAngle60000 == true) frame["rotation"] = shape.Transform.RotationAngle60000 / 60_000d;
        if (shape.Transform?.HasFlipHorizontal == true) frame["flipH"] = shape.Transform.FlipHorizontal;
        if (shape.Transform?.HasFlipVertical == true) frame["flipV"] = shape.Transform.FlipVertical;
        return frame;
    }

    private static JsonObject ImageFrame(PresentationImage image)
    {
        var frame = Frame(image.LeftEmu, image.TopEmu, image.WidthEmu, image.HeightEmu);
        if (image.Transform?.HasRotationAngle60000 == true) frame["rotation"] = image.Transform.RotationAngle60000 / 60_000d;
        if (image.Transform?.HasFlipHorizontal == true) frame["flipH"] = image.Transform.FlipHorizontal;
        if (image.Transform?.HasFlipVertical == true) frame["flipV"] = image.Transform.FlipVertical;
        return frame;
    }

    private static JsonObject TableFrame(PresentationTable table) => Frame(table.LeftEmu, table.TopEmu, table.WidthEmu, table.HeightEmu);
    private static JsonObject ChartFrame(PresentationChart chart) => Frame(chart.LeftEmu, chart.TopEmu, chart.WidthEmu, chart.HeightEmu);
    private static JsonObject GroupFrame(PresentationGroup group) => Frame(group.LeftEmu, group.TopEmu, group.WidthEmu, group.HeightEmu);

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
        PresentationElement.ContentOneofCase.Group => GroupFrame(element.Group),
        PresentationElement.ContentOneofCase.Opaque => Frame(element.Opaque.LeftEmu, element.Opaque.TopEmu, element.Opaque.WidthEmu, element.Opaque.HeightEmu),
        _ => Frame(0, 0, 1, 1),
    };

    private static JsonObject Frame(long left, long top, long width, long height) => new()
    {
        ["x"] = Points(left),
        ["y"] = Points(top),
        ["width"] = Math.Max(0.001, Points(width)),
        ["height"] = Math.Max(0.001, Points(height)),
    };

    private static JsonObject ConnectorEndpoint(
        string targetId,
        long x,
        long y,
        string pageId,
        ProjectionContext context)
    {
        if (!string.IsNullOrEmpty(targetId) && context.TryElementId(pageId, targetId, out var projected))
            return new JsonObject { ["element"] = projected, ["anchor"] = "auto" };
        return new JsonObject { ["x"] = Points(x), ["y"] = Points(y) };
    }

    private static JsonObject Stroke(
        string rgb,
        long widthEmu,
        string? style,
        string? cap,
        string? join,
        double? opacity)
    {
        var output = new JsonObject
        {
            ["color"] = Color(rgb),
            ["width"] = Math.Max(0, Points(widthEmu)),
        };
        var dash = Dash(style);
        if (dash is not null) output["dash"] = dash;
        if (cap is "flat" or "round" or "square") output["cap"] = cap;
        if (join is "miter" or "round" or "bevel") output["join"] = join;
        if (opacity is not null) output["opacity"] = opacity.Value;
        return output;
    }

    private static JsonObject Shadow(PresentationShadow shadow) => new()
    {
        ["color"] = Color(shadow.ColorRgb),
        ["opacity"] = Unit(shadow.OpacityThousandthPercent),
        ["blur"] = Math.Max(0, Points(shadow.BlurRadiusEmu)),
        ["distance"] = Math.Max(0, Points(shadow.DistanceEmu)),
        ["angle"] = shadow.DirectionAngle60000 / 60_000d,
    };

    private static JsonObject? Accessibility(PresentationNonVisualAccessibility? value)
    {
        if (value is null) return null;
        var output = new JsonObject { ["decorative"] = value.HasDecorative && value.Decorative };
        if (!string.IsNullOrEmpty(value.Title)) output["title"] = value.Title;
        if (!string.IsNullOrEmpty(value.Description)) output["description"] = value.Description;
        return output;
    }

    private static JsonObject? ImageAccessibility(PresentationImage image)
    {
        if (!image.HasAccessibilityDecorative && string.IsNullOrEmpty(image.AccessibilityTitle) && string.IsNullOrEmpty(image.AltText)) return null;
        var output = new JsonObject { ["decorative"] = image.HasAccessibilityDecorative && image.AccessibilityDecorative };
        if (!string.IsNullOrEmpty(image.AccessibilityTitle)) output["title"] = image.AccessibilityTitle;
        if (!string.IsNullOrEmpty(image.AltText)) output["description"] = image.AltText;
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
        SpreadsheetChartType.Combo => "combo",
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
        _ => null,
    };

    private static string Color(string rgb) => $"#{rgb.TrimStart('#').ToUpperInvariant()}";
    private static double Crop(int value) => Math.Clamp(value / 100_000d, 0, 1);
    private static double Unit(uint value) => Math.Clamp(value / 100_000d, 0, 1);
    private static double Points(long emu) => Math.Round(emu / EmuPerPoint, 6, MidpointRounding.AwayFromZero);

    private static string HashOrFallback(string? hash, byte[] fallback) =>
        hash is { Length: 64 } && hash.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f')
            ? hash
            : Sha256(fallback);

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
        using var document = JsonDocument.Parse(node.ToJsonString());
        return PpjCanonicalJson.Write(document.RootElement);
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    [GeneratedRegex("[^A-Za-z0-9._:-]+")]
    private static partial Regex InvalidIdCharacters();

    private sealed record CapabilitySpec(string Operation, IReadOnlyList<string> Fields);

    private sealed class ProjectionContext
    {
        private readonly IReadOnlyDictionary<string, Asset> sourceAssets;
        private readonly Dictionary<string, string> pageIds = new(StringComparer.Ordinal);
        private readonly Dictionary<(string Page, string Element), string> elementIds = new();
        private readonly HashSet<string> usedIds = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdBySourceId = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> assetIdByHash = new(StringComparer.Ordinal);
        private readonly JsonArray programAssets = new();
        private readonly List<Asset> resultAssets = [];
        private readonly JsonArray nodes = new();

        internal ProjectionContext(
            string sourceSha256,
            string revision,
            string assetRoot,
            IEnumerable<Asset> assets)
        {
            SourceSha256 = sourceSha256;
            Revision = revision;
            AssetRoot = assetRoot;
            sourceAssets = assets.ToDictionary(asset => asset.Id, StringComparer.Ordinal);
        }

        internal string SourceSha256 { get; }
        internal string Revision { get; }
        internal string AssetRoot { get; }
        internal int VisibleObjectCount { get; private set; }
        internal JsonArray ProgramAssets => programAssets;
        internal IReadOnlyList<Asset> ResultAssets => resultAssets;

        internal string RegisterPage(string sourceId)
        {
            var id = UniqueId($"page-{NormalizeId(sourceId, "slide")}");
            pageIds[sourceId] = id;
            return id;
        }

        internal string RegisterElement(string pageId, string sourceId)
        {
            var id = UniqueId($"{pageId}-{NormalizeId(sourceId, "element")}");
            elementIds[(pageId, sourceId)] = id;
            return id;
        }

        internal string PageId(string sourceId) => pageIds[sourceId];
        internal bool TryPageId(string sourceId, out string id) => pageIds.TryGetValue(sourceId, out id!);
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
            if (!sourceAssets.TryGetValue(sourceId, out var source) || source.Data.IsEmpty) return false;
            var hash = HashOrFallback(source.Sha256, source.Data.ToByteArray());
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
                ["id"] = programAssetId,
                ["uri"] = $"{AssetRoot}/{fileName}",
                ["mimeType"] = source.ContentType,
                ["sha256"] = hash,
                ["rights"] = new JsonObject { ["status"] = "user-provided" },
                ["accessibility"] = new JsonObject { ["decorative"] = false, ["description"] = "Imported source media." },
            });
            var materialized = source.Clone();
            materialized.Id = programAssetId;
            materialized.FileName = fileName;
            materialized.Sha256 = hash;
            resultAssets.Add(materialized);
            return true;
        }

        internal void RecordNode(string pageId, string id, string type, JsonObject nativeRef)
        {
            VisibleObjectCount++;
            nodes.Add(new JsonObject
            {
                ["id"] = id,
                ["page"] = pageId,
                ["type"] = type,
                ["handle"] = nativeRef["handle"]!.GetValue<string>(),
                ["objectHash"] = nativeRef["objectHash"]!.GetValue<string>(),
                ["capabilitySetSha256"] = nativeRef["capabilitySetSha256"]!.GetValue<string>(),
            });
        }

        internal JsonObject BuildNodeMap() => new()
        {
            ["schema"] = "office-kit/ppj-node-map/v1",
            ["sourceSha256"] = SourceSha256,
            ["revision"] = Revision,
            ["nodes"] = nodes.DeepClone(),
        };

        private static string Extension(string contentType, string fileName)
        {
            var fromType = contentType.ToLowerInvariant() switch
            {
                "image/png" => "png",
                "image/jpeg" => "jpg",
                "image/gif" => "gif",
                "image/svg+xml" => "svg",
                _ => string.Empty,
            };
            if (fromType.Length > 0) return fromType;
            var extension = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            return Regex.IsMatch(extension, "^[a-z0-9]{1,8}$") ? extension : "bin";
        }
    }
}
