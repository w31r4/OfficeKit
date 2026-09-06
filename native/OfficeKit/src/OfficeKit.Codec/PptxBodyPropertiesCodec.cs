using DocumentFormat.OpenXml;
using System.Globalization;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

// Owns the bounded a:bodyPr layout subset. Unmodeled attributes and children
// remain on the source element. Canonical normAutofit percentages are modeled
// with exact attribute presence; other AutoFit markup stays source-bound.
internal static class PptxBodyPropertiesCodec
{
    private const int MaxRotationAngle60000 = 21_600_000;
    private const int MinFontScale1000 = 1_000;
    private const int MaxFontScale1000 = 100_000;
    private const int MinLineSpacingReduction1000 = 0;
    private const int MaxLineSpacingReduction1000 = 13_200_000;
    private const int MaxTextWarpAdjustments = 256;
    private const long MinFlatTextZ = int.MinValue;
    private const long MaxFlatTextZ = int.MaxValue;
    private static readonly HashSet<string> TextWarpPresets = new(StringComparer.Ordinal)
    {
        "textNoShape", "textPlain", "textStop", "textTriangle", "textTriangleInverted",
        "textChevron", "textChevronInverted", "textRingInside", "textRingOutside",
        "textArchUp", "textArchDown", "textCircle", "textButton", "textArchUpPour",
        "textArchDownPour", "textCirclePour", "textButtonPour", "textCurveUp",
        "textCurveDown", "textCanUp", "textCanDown", "textWave1", "textWave2",
        "textDoubleWave1", "textWave4", "textInflate", "textDeflate", "textInflateBottom",
        "textDeflateBottom", "textInflateTop", "textDeflateTop", "textDeflateInflate",
        "textDeflateInflateDeflate", "textFadeRight", "textFadeLeft", "textFadeUp",
        "textFadeDown", "textSlantUp", "textSlantDown", "textCascadeUp", "textCascadeDown",
    };

    internal static void Read(PresentationTextBody target, P.TextBody source)
    {
        var native = source.Elements<A.BodyProperties>().FirstOrDefault();
        if (native is null) return;
        var modeled = new PresentationTextBodyProperties();
        ReadInset(native.LeftInset?.Value, value => modeled.LeftInsetEmu = value);
        ReadInset(native.TopInset?.Value, value => modeled.TopInsetEmu = value);
        ReadInset(native.RightInset?.Value, value => modeled.RightInsetEmu = value);
        ReadInset(native.BottomInset?.Value, value => modeled.BottomInsetEmu = value);
        if (AnchorName(native.Anchor?.Value) is { Length: > 0 } anchor) modeled.VerticalAnchor = anchor;
        if (WrapName(native.Wrap?.Value) is { Length: > 0 } wrap) modeled.Wrap = wrap;
        var autoFit = native.ChildElements.Where(IsAutoFitChoice).ToArray();
        if (autoFit.Length == 1 && SupportsAutoFitChoice(autoFit[0])) ReadAutoFit(modeled, autoFit[0]);
        if (native.Rotation?.Value is >= -MaxRotationAngle60000 and <= MaxRotationAngle60000) modeled.RotationAngle60000 = native.Rotation.Value;
        if (VerticalTextName(native.Vertical?.Value) is { Length: > 0 } verticalText) modeled.VerticalTextMode = verticalText;
        if (VerticalOverflowName(native.VerticalOverflow?.Value) is { Length: > 0 } verticalOverflow) modeled.VerticalOverflowMode = verticalOverflow;
        if (HorizontalOverflowName(native.HorizontalOverflow?.Value) is { Length: > 0 } horizontalOverflow) modeled.HorizontalOverflowMode = horizontalOverflow;
        if (native.ColumnCount?.Value is >= 1 and <= 16) modeled.Columns = checked((uint)native.ColumnCount.Value);
        if (native.ColumnSpacing?.Value is >= 0) modeled.ColumnSpacingEmu = native.ColumnSpacing.Value;
        if (native.RightToLeftColumns?.Value is { } rightToLeft) modeled.RightToLeftColumns = rightToLeft;
        if (native.UpRight?.Value is { } upright) modeled.Upright = upright;
        if (native.AnchorCenter?.Value is { } anchorCenter) modeled.AnchorCenter = anchorCenter;
        if (native.ForceAntiAlias?.Value is { } forceAntiAlias) modeled.ForceAntiAlias = forceAntiAlias;
        if (native.UseParagraphSpacing?.Value is { } spaceFirstLastParagraph) modeled.SpaceFirstLastParagraph = spaceFirstLastParagraph;
        if (native.CompatibleLineSpacing?.Value is { } compatibleLineSpacing) modeled.CompatibleLineSpacing = compatibleLineSpacing;
        if (native.FromWordArt?.Value is { } fromWordArt) modeled.FromWordArt = fromWordArt;
        if (TryReadTextWarp(native, out var textWarpPreset, out var textWarpAdjustments))
        {
            modeled.TextWarpPreset = textWarpPreset;
            modeled.TextWarpAdjustments.Add(textWarpAdjustments);
        }
        if (TryReadFlatTextZ(native, out var flatTextZ)) modeled.FlatTextZ = flatTextZ;
        if (HasModeledProperties(modeled)) target.BodyProperties = modeled;
    }

    internal static bool Supports(P.TextBody? source)
    {
        if (source is null) return true;
        var bodies = source.Elements<A.BodyProperties>().ToArray();
        return bodies.Length <= 1 && (bodies.Length == 0 || bodies[0].ChildElements.Count(IsAutoFitChoice) <= 1);
    }

    internal static void Validate(PresentationTextBody source)
    {
        if (source.BodyProperties is null) return;
        var properties = source.BodyProperties;
        ValidateInset(properties.LeftInsetCase, properties.LeftInsetEmu, PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, PresentationTextBodyProperties.LeftInsetOneofCase.NoLeftInset, properties.NoLeftInset, "left");
        ValidateInset(properties.TopInsetCase, properties.TopInsetEmu, PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, PresentationTextBodyProperties.TopInsetOneofCase.NoTopInset, properties.NoTopInset, "top");
        ValidateInset(properties.RightInsetCase, properties.RightInsetEmu, PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, PresentationTextBodyProperties.RightInsetOneofCase.NoRightInset, properties.NoRightInset, "right");
        ValidateInset(properties.BottomInsetCase, properties.BottomInsetEmu, PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, PresentationTextBodyProperties.BottomInsetOneofCase.NoBottomInset, properties.NoBottomInset, "bottom");
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) _ = ParseAnchor(properties.VerticalAnchor);
        else if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.NoVerticalAnchor && !properties.NoVerticalAnchor) throw Invalid("Presentation no_vertical_anchor must be true when selected.");
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) _ = ParseWrap(properties.Wrap);
        else if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.NoWrap && !properties.NoWrap) throw Invalid("Presentation no_wrap must be true when selected.");
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode) _ = ParseAutoFit(properties.AutoFitMode);
        else if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.NoAutoFitMode && !properties.NoAutoFitMode) throw Invalid("Presentation no_auto_fit_mode must be true when selected.");
        if (properties.HasTextWarpPreset) _ = ParseTextWarpPreset(properties.TextWarpPreset);
        ValidateTextWarpAdjustments(properties);
        if (properties.HasFlatTextZ && (properties.FlatTextZ < MinFlatTextZ || properties.FlatTextZ > MaxFlatTextZ))
            throw Invalid("Presentation flat-text z coordinate must fit the bounded signed 32-bit range.");
        ValidateNormalAutoFit(properties);
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000 && Math.Abs((long)properties.RotationAngle60000) > MaxRotationAngle60000) throw Invalid("Presentation text body rotation must be between -360 and 360 degrees.");
        else if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.NoRotation && !properties.NoRotation) throw Invalid("Presentation no_rotation must be true when selected.");
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) _ = ParseVerticalText(properties.VerticalTextMode);
        else if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.NoVerticalTextMode && !properties.NoVerticalTextMode) throw Invalid("Presentation no_vertical_text_mode must be true when selected.");
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) _ = ParseVerticalOverflow(properties.VerticalOverflowMode);
        else if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.NoVerticalOverflowMode && !properties.NoVerticalOverflowMode) throw Invalid("Presentation no_vertical_overflow_mode must be true when selected.");
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) _ = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        else if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.NoHorizontalOverflowMode && !properties.NoHorizontalOverflowMode) throw Invalid("Presentation no_horizontal_overflow_mode must be true when selected.");
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns && (properties.Columns < 1 || properties.Columns > 16)) throw Invalid("Presentation text body column count must be from 1 through 16.");
        else if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.NoColumns && !properties.NoColumns) throw Invalid("Presentation no_columns must be true when selected.");
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu && (properties.ColumnSpacingEmu < 0 || properties.ColumnSpacingEmu > int.MaxValue)) throw Invalid("Presentation text body column spacing must fit the non-negative signed 32-bit EMU range.");
        else if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.NoColumnSpacing && !properties.NoColumnSpacing) throw Invalid("Presentation no_column_spacing must be true when selected.");
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.NoColumnDirection && !properties.NoColumnDirection) throw Invalid("Presentation no_column_direction must be true when selected.");
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.NoUpright && !properties.NoUpright) throw Invalid("Presentation no_upright must be true when selected.");
    }

    internal static bool HasModeledProperties(PresentationTextBodyProperties? source) => source is not null &&
        (source.LeftInsetCase != PresentationTextBodyProperties.LeftInsetOneofCase.None ||
         source.TopInsetCase != PresentationTextBodyProperties.TopInsetOneofCase.None ||
         source.RightInsetCase != PresentationTextBodyProperties.RightInsetOneofCase.None ||
         source.BottomInsetCase != PresentationTextBodyProperties.BottomInsetOneofCase.None ||
         source.AnchorCase != PresentationTextBodyProperties.AnchorOneofCase.None ||
         source.WrappingCase != PresentationTextBodyProperties.WrappingOneofCase.None ||
         source.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.None ||
         source.RotationCase != PresentationTextBodyProperties.RotationOneofCase.None ||
         source.VerticalTextCase != PresentationTextBodyProperties.VerticalTextOneofCase.None ||
         source.VerticalOverflowCase != PresentationTextBodyProperties.VerticalOverflowOneofCase.None ||
         source.HorizontalOverflowCase != PresentationTextBodyProperties.HorizontalOverflowOneofCase.None ||
         source.ColumnCountCase != PresentationTextBodyProperties.ColumnCountOneofCase.None ||
         source.ColumnSpacingCase != PresentationTextBodyProperties.ColumnSpacingOneofCase.None ||
         source.ColumnDirectionCase != PresentationTextBodyProperties.ColumnDirectionOneofCase.None ||
         source.UprightTextCase != PresentationTextBodyProperties.UprightTextOneofCase.None ||
         source.HasAnchorCenter ||
         source.HasForceAntiAlias ||
         source.HasSpaceFirstLastParagraph ||
         source.HasCompatibleLineSpacing ||
         source.HasFromWordArt ||
         source.HasTextWarpPreset ||
         source.TextWarpAdjustments.Count > 0 ||
         source.HasFlatTextZ);

    // A source-bound text-body style may expose only direct bodyPr leaves with
    // a stable PPJ textBoxStyle spelling.  The bounded profile includes the
    // scalar rotation, overflow and upright attributes and canonical
    // normAutofit percentages.  Other bodyPr children remain source-owned.
    internal static bool SupportsBoundedDirectLayout(PresentationTextBodyProperties? source)
    {
        if (source is null) return true;
        return (source.LeftInsetCase is PresentationTextBodyProperties.LeftInsetOneofCase.None or PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) &&
            (source.TopInsetCase is PresentationTextBodyProperties.TopInsetOneofCase.None or PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) &&
            (source.RightInsetCase is PresentationTextBodyProperties.RightInsetOneofCase.None or PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) &&
            (source.BottomInsetCase is PresentationTextBodyProperties.BottomInsetOneofCase.None or PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) &&
            (source.AnchorCase is PresentationTextBodyProperties.AnchorOneofCase.None or PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) &&
            (source.WrappingCase is PresentationTextBodyProperties.WrappingOneofCase.None or PresentationTextBodyProperties.WrappingOneofCase.Wrap) &&
            (source.AutoFitCase is PresentationTextBodyProperties.AutoFitOneofCase.None or PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode) &&
            (source.VerticalTextCase is PresentationTextBodyProperties.VerticalTextOneofCase.None or PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) &&
            (source.ColumnCountCase is PresentationTextBodyProperties.ColumnCountOneofCase.None or PresentationTextBodyProperties.ColumnCountOneofCase.Columns) &&
            (source.ColumnSpacingCase is PresentationTextBodyProperties.ColumnSpacingOneofCase.None or PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu) &&
            (source.ColumnDirectionCase is PresentationTextBodyProperties.ColumnDirectionOneofCase.None or PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns) &&
            (source.RotationCase is PresentationTextBodyProperties.RotationOneofCase.None or PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000) &&
            (source.VerticalOverflowCase is PresentationTextBodyProperties.VerticalOverflowOneofCase.None or PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) &&
            (source.HorizontalOverflowCase is PresentationTextBodyProperties.HorizontalOverflowOneofCase.None or PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) &&
            (source.UprightTextCase is PresentationTextBodyProperties.UprightTextOneofCase.None or PresentationTextBodyProperties.UprightTextOneofCase.Upright) &&
            SupportsBoundedNormalAutoFit(source);
    }

    private static bool SupportsBoundedNormalAutoFit(PresentationTextBodyProperties source)
    {
        if (source.NormalAutoFit is not { } normal) return true;
        if (source.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode || source.AutoFitMode != "shrinkText")
            return false;
        if (normal.FontScaleCase is not (PresentationNormalAutoFit.FontScaleOneofCase.None or PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000) ||
            normal.LineSpacingReductionCase is not (PresentationNormalAutoFit.LineSpacingReductionOneofCase.None or PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000))
            return false;
        if (normal.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000 &&
            (normal.FontScale1000 < MinFontScale1000 || normal.FontScale1000 > MaxFontScale1000))
            return false;
        if (normal.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000 &&
            (normal.LineSpacingReduction1000 < MinLineSpacingReduction1000 || normal.LineSpacingReduction1000 > MaxLineSpacingReduction1000))
            return false;
        return true;
    }

    internal static void Build(A.BodyProperties target, PresentationTextBody source)
    {
        if (source.BodyProperties is not { } properties) return;
        ApplyInsets(target, properties);
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) target.Anchor = ParseAnchor(properties.VerticalAnchor);
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) target.Wrap = ParseWrap(properties.Wrap);
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000) target.Rotation = properties.RotationAngle60000;
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) target.Vertical = ParseVerticalText(properties.VerticalTextMode);
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) target.VerticalOverflow = ParseVerticalOverflow(properties.VerticalOverflowMode);
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) target.HorizontalOverflow = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns) target.ColumnCount = checked((int)properties.Columns);
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu) target.ColumnSpacing = checked((int)properties.ColumnSpacingEmu);
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns) target.RightToLeftColumns = properties.RightToLeftColumns;
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright) target.UpRight = properties.Upright;
        if (properties.HasAnchorCenter) target.AnchorCenter = properties.AnchorCenter;
        if (properties.HasForceAntiAlias) target.ForceAntiAlias = properties.ForceAntiAlias;
        if (properties.HasSpaceFirstLastParagraph) target.UseParagraphSpacing = properties.SpaceFirstLastParagraph;
        if (properties.HasCompatibleLineSpacing) target.CompatibleLineSpacing = properties.CompatibleLineSpacing;
        if (properties.HasFromWordArt) target.FromWordArt = properties.FromWordArt;
        if (properties.HasTextWarpPreset)
        {
            var textWarp = new A.PresetTextWarp
            {
                Preset = new A.TextShapeValues(ParseTextWarpPreset(properties.TextWarpPreset)),
            };
            ApplyTextWarpAdjustments(textWarp, properties.TextWarpAdjustments);
            target.AddChild(textWarp, true);
        }
        if (properties.HasFlatTextZ)
            target.AddChild(new A.FlatText { Z = properties.FlatTextZ }, true);
        if (properties.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode) target.AddChild(CreateAutoFit(properties.AutoFitMode, properties.NormalAutoFit), true);
    }

    internal static void Apply(P.TextBody target, PresentationTextBody source)
    {
        var bodies = target.Elements<A.BodyProperties>().ToArray();
        if (bodies.Length > 1) throw Unsupported("Source-preserving PPTX export cannot edit duplicate text body properties.");
        var native = bodies.FirstOrDefault();
        if (native is null)
        {
            if (!HasModeledProperties(source.BodyProperties)) return;
            native = new A.BodyProperties();
            target.PrependChild(native);
        }
        var properties = source.BodyProperties;
        if (properties is null) return;
        ApplyInset(properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu, properties.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.NoLeftInset, properties.LeftInsetEmu, value => native.LeftInset = value, () => native.LeftInset = null);
        ApplyInset(properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu, properties.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.NoTopInset, properties.TopInsetEmu, value => native.TopInset = value, () => native.TopInset = null);
        ApplyInset(properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu, properties.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.NoRightInset, properties.RightInsetEmu, value => native.RightInset = value, () => native.RightInset = null);
        ApplyInset(properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu, properties.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.NoBottomInset, properties.BottomInsetEmu, value => native.BottomInset = value, () => native.BottomInset = null);
        if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.VerticalAnchor) native.Anchor = ParseAnchor(properties.VerticalAnchor);
        else if (properties.AnchorCase == PresentationTextBodyProperties.AnchorOneofCase.NoVerticalAnchor) native.Anchor = null;
        if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.Wrap) native.Wrap = ParseWrap(properties.Wrap);
        else if (properties.WrappingCase == PresentationTextBodyProperties.WrappingOneofCase.NoWrap) native.Wrap = null;
        if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.RotationAngle60000) native.Rotation = properties.RotationAngle60000;
        else if (properties.RotationCase == PresentationTextBodyProperties.RotationOneofCase.NoRotation) native.Rotation = null;
        if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.VerticalTextMode) native.Vertical = ParseVerticalText(properties.VerticalTextMode);
        else if (properties.VerticalTextCase == PresentationTextBodyProperties.VerticalTextOneofCase.NoVerticalTextMode) native.Vertical = null;
        if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.VerticalOverflowMode) native.VerticalOverflow = ParseVerticalOverflow(properties.VerticalOverflowMode);
        else if (properties.VerticalOverflowCase == PresentationTextBodyProperties.VerticalOverflowOneofCase.NoVerticalOverflowMode) native.VerticalOverflow = null;
        if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.HorizontalOverflowMode) native.HorizontalOverflow = ParseHorizontalOverflow(properties.HorizontalOverflowMode);
        else if (properties.HorizontalOverflowCase == PresentationTextBodyProperties.HorizontalOverflowOneofCase.NoHorizontalOverflowMode) native.HorizontalOverflow = null;
        if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.Columns) native.ColumnCount = checked((int)properties.Columns);
        else if (properties.ColumnCountCase == PresentationTextBodyProperties.ColumnCountOneofCase.NoColumns) native.ColumnCount = null;
        if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.ColumnSpacingEmu) native.ColumnSpacing = checked((int)properties.ColumnSpacingEmu);
        else if (properties.ColumnSpacingCase == PresentationTextBodyProperties.ColumnSpacingOneofCase.NoColumnSpacing) native.ColumnSpacing = null;
        if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.RightToLeftColumns) native.RightToLeftColumns = properties.RightToLeftColumns;
        else if (properties.ColumnDirectionCase == PresentationTextBodyProperties.ColumnDirectionOneofCase.NoColumnDirection) native.RightToLeftColumns = null;
        if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.Upright) native.UpRight = properties.Upright;
        else if (properties.UprightTextCase == PresentationTextBodyProperties.UprightTextOneofCase.NoUpright) native.UpRight = null;
        if (properties.HasAnchorCenter) native.AnchorCenter = properties.AnchorCenter;
        if (properties.HasForceAntiAlias) native.ForceAntiAlias = properties.ForceAntiAlias;
        if (properties.HasSpaceFirstLastParagraph) native.UseParagraphSpacing = properties.SpaceFirstLastParagraph;
        if (properties.HasCompatibleLineSpacing) native.CompatibleLineSpacing = properties.CompatibleLineSpacing;
        if (properties.HasFromWordArt) native.FromWordArt = properties.FromWordArt;
        if (properties.HasTextWarpPreset)
        {
            var choices = native.ChildElements.OfType<A.PresetTextWarp>().ToArray();
            if (choices.Length > 1)
                throw Unsupported("Source-preserving PPTX export cannot edit duplicate text-warp presets.");
            if (choices.Length == 1)
            {
                if (!TryReadTextWarp(choices[0], out _, out _))
                    throw Unsupported("Source-preserving PPTX export cannot replace noncanonical text-warp preset markup.");
                choices[0].Preset = new A.TextShapeValues(ParseTextWarpPreset(properties.TextWarpPreset));
                ApplyTextWarpAdjustments(choices[0], properties.TextWarpAdjustments);
            }
            else
            {
                var textWarp = new A.PresetTextWarp
                {
                    Preset = new A.TextShapeValues(ParseTextWarpPreset(properties.TextWarpPreset)),
                };
                ApplyTextWarpAdjustments(textWarp, properties.TextWarpAdjustments);
                native.AddChild(textWarp, true);
            }
        }
        if (properties.HasFlatTextZ)
        {
            var flatTexts = native.ChildElements.OfType<A.FlatText>().ToArray();
            if (flatTexts.Length > 1)
                throw Unsupported("Source-preserving PPTX export cannot edit duplicate flat-text children.");
            if (flatTexts.Length == 1)
            {
                if (!TryReadFlatTextZ(flatTexts[0], out _))
                    throw Unsupported("Source-preserving PPTX export cannot replace noncanonical flat-text markup.");
                flatTexts[0].Z = properties.FlatTextZ;
            }
            else
            {
                native.AddChild(new A.FlatText { Z = properties.FlatTextZ }, true);
            }
        }
        ApplyAutoFit(native, properties);
    }

    internal static void Scrub(P.TextBody? source)
    {
        foreach (var native in source?.Elements<A.BodyProperties>() ?? [])
        {
            if (native.LeftInset?.Value is >= 0) native.LeftInset = null;
            if (native.TopInset?.Value is >= 0) native.TopInset = null;
            if (native.RightInset?.Value is >= 0) native.RightInset = null;
            if (native.BottomInset?.Value is >= 0) native.BottomInset = null;
            if (AnchorName(native.Anchor?.Value).Length > 0) native.Anchor = null;
            if (WrapName(native.Wrap?.Value).Length > 0) native.Wrap = null;
            if (native.Rotation?.Value is >= -MaxRotationAngle60000 and <= MaxRotationAngle60000) native.Rotation = null;
            if (VerticalTextName(native.Vertical?.Value).Length > 0) native.Vertical = null;
            if (VerticalOverflowName(native.VerticalOverflow?.Value).Length > 0) native.VerticalOverflow = null;
            if (HorizontalOverflowName(native.HorizontalOverflow?.Value).Length > 0) native.HorizontalOverflow = null;
            if (native.ColumnCount?.Value is >= 1 and <= 16) native.ColumnCount = null;
            if (native.ColumnSpacing?.Value is >= 0) native.ColumnSpacing = null;
            if (native.RightToLeftColumns is not null) native.RightToLeftColumns = null;
            if (native.UpRight is not null) native.UpRight = null;
            if (native.AnchorCenter is not null) native.AnchorCenter = null;
            if (native.ForceAntiAlias is not null) native.ForceAntiAlias = null;
            if (native.UseParagraphSpacing is not null) native.UseParagraphSpacing = null;
            if (native.CompatibleLineSpacing is not null) native.CompatibleLineSpacing = null;
            if (native.FromWordArt is not null) native.FromWordArt = null;
            foreach (var textWarp in native.ChildElements.OfType<A.PresetTextWarp>().Where(item => TryReadTextWarpPreset(item, out _)).ToArray()) textWarp.Remove();
            foreach (var flatText in native.ChildElements.OfType<A.FlatText>().Where(item => TryReadFlatTextZ(item, out _)).ToArray()) flatText.Remove();
            foreach (var autoFit in native.ChildElements.Where(child => IsAutoFitChoice(child) && SupportsAutoFitChoice(child)).ToArray()) autoFit.Remove();
        }
    }

    internal static bool TryReadTextWarpPreset(A.BodyProperties source, out string preset)
    {
        return TryReadTextWarp(source, out preset, out _);
    }

    internal static bool TryReadTextWarpPreset(A.PresetTextWarp source, out string preset)
    {
        return TryReadTextWarp(source, out preset, out _);
    }

    internal static bool TryReadTextWarp(
        A.BodyProperties source,
        out string preset,
        out IReadOnlyList<PresentationTextWarpAdjustment> adjustments)
    {
        preset = string.Empty;
        adjustments = [];
        var choices = source.ChildElements.OfType<A.PresetTextWarp>().ToArray();
        return choices.Length == 1 && TryReadTextWarp(choices[0], out preset, out adjustments);
    }

    internal static bool TryReadTextWarp(
        A.PresetTextWarp source,
        out string preset,
        out IReadOnlyList<PresentationTextWarpAdjustment> adjustments)
    {
        preset = string.Empty;
        adjustments = [];
        var attributes = source.GetAttributes();
        if (attributes.Count != 1 || attributes[0].NamespaceUri.Length != 0 || attributes[0].LocalName != "prst") return false;
        if (attributes[0].Value is not { Length: > 0 } value || !TextWarpPresets.Contains(value)) return false;
        preset = value;
        if (source.ChildElements.Count == 0) return true;
        if (source.ChildElements.Count != 1 || source.FirstChild is not A.AdjustValueList list || list.GetAttributes().Count != 0)
        {
            preset = string.Empty;
            return false;
        }
        var nativeGuides = list.Elements<A.ShapeGuide>().ToArray();
        if (nativeGuides.Length == 0 || nativeGuides.Length > MaxTextWarpAdjustments ||
            list.ChildElements.Count != nativeGuides.Length)
        {
            preset = string.Empty;
            return false;
        }
        var names = new HashSet<string>(StringComparer.Ordinal);
        var modeled = new List<PresentationTextWarpAdjustment>(nativeGuides.Length);
        foreach (var native in nativeGuides)
        {
            var guideAttributes = native.GetAttributes();
            if (native.ChildElements.Count != 0 || guideAttributes.Count != 2 ||
                guideAttributes.Any(attribute => attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("name" or "fmla")) ||
                native.Name?.Value is not { Length: > 0 } name || name.Length > 256 || name.Any(char.IsControl) || !names.Add(name) ||
                native.Formula?.Value is not { Length: > 0 } formula ||
                !TryLiteralTextWarpAdjustment(formula, out var adjustment))
            {
                preset = string.Empty;
                adjustments = [];
                return false;
            }
            modeled.Add(new PresentationTextWarpAdjustment { Name = name, Value = adjustment });
        }
        adjustments = modeled;
        return true;
    }

    internal static bool TryLiteralTextWarpAdjustment(string formula, out int value)
    {
        value = 0;
        var tokens = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 2 || tokens[0] != "val" ||
            !int.TryParse(tokens[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed.ToString(CultureInfo.InvariantCulture) != tokens[1])
            return false;
        value = parsed;
        return true;
    }

    internal static int ParseTextWarpAdjustment(string value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) &&
        parsed.ToString(CultureInfo.InvariantCulture) == value
            ? parsed
            : throw Invalid($"Unsupported Presentation text-warp adjustment {value}.");

    internal static bool TryReadFlatTextZ(A.BodyProperties source, out long value)
    {
        value = 0;
        var flatTexts = source.ChildElements.OfType<A.FlatText>().ToArray();
        return flatTexts.Length == 1 && TryReadFlatTextZ(flatTexts[0], out value);
    }

    internal static bool TryReadFlatTextZ(A.FlatText source, out long value)
    {
        value = 0;
        var attributes = source.GetAttributes();
        if (source.ChildElements.Count != 0 || attributes.Count != 1 ||
            attributes[0].NamespaceUri.Length != 0 || attributes[0].LocalName != "z" ||
            !long.TryParse(attributes[0].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ||
            parsed.ToString(CultureInfo.InvariantCulture) != attributes[0].Value ||
            parsed < MinFlatTextZ || parsed > MaxFlatTextZ)
            return false;
        value = parsed;
        return true;
    }

    internal static long ParseFlatTextZ(string value) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) &&
        parsed.ToString(CultureInfo.InvariantCulture) == value &&
        parsed >= MinFlatTextZ && parsed <= MaxFlatTextZ
            ? parsed
            : throw Invalid($"Unsupported Presentation flat-text z coordinate {value}.");

    private static void ValidateTextWarpAdjustments(PresentationTextBodyProperties source)
    {
        if (source.TextWarpAdjustments.Count == 0) return;
        if (!source.HasTextWarpPreset)
            throw Invalid("Presentation text-warp adjustments require text_warp_preset.");
        if (source.TextWarpAdjustments.Count > MaxTextWarpAdjustments)
            throw Invalid($"Presentation text-warp adjustments cannot exceed {MaxTextWarpAdjustments} entries.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var adjustment in source.TextWarpAdjustments)
        {
            if (adjustment.Name.Length == 0 || adjustment.Name.Length > 256 || adjustment.Name.Any(char.IsControl) || !names.Add(adjustment.Name))
                throw Invalid("Presentation text-warp adjustment names must be unique, non-empty, and free of control characters.");
        }
    }

    private static void ApplyTextWarpAdjustments(
        A.PresetTextWarp target,
        IEnumerable<PresentationTextWarpAdjustment> source)
    {
        var adjustments = source.ToArray();
        if (adjustments.Length == 0)
        {
            target.GetFirstChild<A.AdjustValueList>()?.Remove();
            return;
        }
        target.GetFirstChild<A.AdjustValueList>()?.Remove();
        target.AppendChild(new A.AdjustValueList(adjustments.Select(adjustment =>
            new A.ShapeGuide
            {
                Name = adjustment.Name,
                Formula = $"val {adjustment.Value.ToString(CultureInfo.InvariantCulture)}",
            })));
    }

    internal static string ParseTextWarpPreset(string value) => TextWarpPresets.Contains(value)
        ? value
        : throw Invalid($"Unsupported Presentation text-warp preset {value}.");

    private static void ReadInset(int? value, Action<long> assign)
    {
        if (value is >= 0) assign(value.Value);
    }

    private static void ValidateInset<TCase>(TCase actualCase, long value, TCase valueCase, TCase noCase, bool noValue, string name) where TCase : struct, Enum
    {
        if (EqualityComparer<TCase>.Default.Equals(actualCase, valueCase) && (value < 0 || value > int.MaxValue)) throw Invalid($"Presentation {name} text inset must fit the non-negative signed 32-bit EMU range.");
        if (EqualityComparer<TCase>.Default.Equals(actualCase, noCase) && !noValue) throw Invalid($"Presentation no_{name}_inset must be true when selected.");
    }

    private static void ApplyInsets(A.BodyProperties target, PresentationTextBodyProperties source)
    {
        if (source.LeftInsetCase == PresentationTextBodyProperties.LeftInsetOneofCase.LeftInsetEmu) target.LeftInset = checked((int)source.LeftInsetEmu);
        if (source.TopInsetCase == PresentationTextBodyProperties.TopInsetOneofCase.TopInsetEmu) target.TopInset = checked((int)source.TopInsetEmu);
        if (source.RightInsetCase == PresentationTextBodyProperties.RightInsetOneofCase.RightInsetEmu) target.RightInset = checked((int)source.RightInsetEmu);
        if (source.BottomInsetCase == PresentationTextBodyProperties.BottomInsetOneofCase.BottomInsetEmu) target.BottomInset = checked((int)source.BottomInsetEmu);
    }

    private static void ApplyInset(bool hasValue, bool hasDelete, long value, Action<int> set, Action clear)
    {
        if (hasValue) set(checked((int)value));
        else if (hasDelete) clear();
    }

    private static void ApplyAutoFit(A.BodyProperties target, PresentationTextBodyProperties source)
    {
        if (source.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.None) return;
        var choices = target.ChildElements.Where(IsAutoFitChoice).ToArray();
        if (choices.Length > 1) throw Unsupported("Source-preserving PPTX export cannot replace duplicate AutoFit choices.");
        var current = choices.FirstOrDefault();
        if (current is not null && !SupportsAutoFitChoice(current)) throw Unsupported("Source-preserving PPTX export cannot replace noncanonical AutoFit markup.");
        if (source.AutoFitCase == PresentationTextBodyProperties.AutoFitOneofCase.NoAutoFitMode)
        {
            current?.Remove();
            return;
        }
        var mode = source.AutoFitMode;
        if (current is A.NormalAutoFit normal && mode == "shrinkText")
        {
            ApplyNormalAutoFit(normal, source.NormalAutoFit);
            return;
        }
        if (current is not null && AutoFitName(current) == mode) return;
        current?.Remove();
        target.AddChild(CreateAutoFit(mode, source.NormalAutoFit), true);
    }

    private static bool IsAutoFitChoice(OpenXmlElement child) => child is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit;
    private static bool IsSimple(OpenXmlElement child) => child.GetAttributes().Count == 0 && child.ChildElements.Count == 0;
    private static bool SupportsAutoFitChoice(OpenXmlElement child) => child switch
    {
        A.NoAutoFit or A.ShapeAutoFit => IsSimple(child),
        A.NormalAutoFit normal => TryReadNormalAutoFit(normal, out _, out _),
        _ => false,
    };

    private static void ReadAutoFit(PresentationTextBodyProperties target, OpenXmlElement source)
    {
        target.AutoFitMode = AutoFitName(source);
        if (source is not A.NormalAutoFit normal || !TryReadNormalAutoFit(normal, out var fontScale, out var lineSpacingReduction)) return;
        if (fontScale is null && lineSpacingReduction is null) return;
        target.NormalAutoFit = new PresentationNormalAutoFit();
        if (fontScale is not null) target.NormalAutoFit.FontScale1000 = fontScale.Value;
        if (lineSpacingReduction is not null) target.NormalAutoFit.LineSpacingReduction1000 = lineSpacingReduction.Value;
    }

    private static bool TryReadNormalAutoFit(A.NormalAutoFit source, out int? fontScale, out int? lineSpacingReduction)
    {
        fontScale = null;
        lineSpacingReduction = null;
        if (source.ChildElements.Count != 0) return false;
        foreach (var attribute in source.GetAttributes())
        {
            if (attribute.NamespaceUri.Length != 0) return false;
            if (!int.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value)) return false;
            if (attribute.LocalName == "fontScale" && value is >= MinFontScale1000 and <= MaxFontScale1000) fontScale = value;
            else if (attribute.LocalName == "lnSpcReduction" && value is >= MinLineSpacingReduction1000 and <= MaxLineSpacingReduction1000) lineSpacingReduction = value;
            else return false;
        }
        return true;
    }
    private static string AutoFitName(OpenXmlElement child) => child switch
    {
        A.NoAutoFit => "none",
        A.NormalAutoFit => "shrinkText",
        A.ShapeAutoFit => "resizeShape",
        _ => string.Empty,
    };

    private static OpenXmlElement CreateAutoFit(string value, PresentationNormalAutoFit? profile) => ParseAutoFit(value) switch
    {
        "none" => new A.NoAutoFit(),
        "shrinkText" => CreateNormalAutoFit(profile),
        "resizeShape" => new A.ShapeAutoFit(),
        _ => throw Invalid($"Unsupported Presentation AutoFit mode {value}."),
    };

    private static A.NormalAutoFit CreateNormalAutoFit(PresentationNormalAutoFit? source)
    {
        var target = new A.NormalAutoFit();
        ApplyNormalAutoFit(target, source);
        return target;
    }

    private static void ApplyNormalAutoFit(A.NormalAutoFit target, PresentationNormalAutoFit? source)
    {
        if (source is null) return;
        if (source.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000) target.FontScale = source.FontScale1000;
        else if (source.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.NoFontScale) target.FontScale = null;
        if (source.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000) target.LineSpaceReduction = source.LineSpacingReduction1000;
        else if (source.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.NoLineSpacingReduction) target.LineSpaceReduction = null;
    }

    private static void ValidateNormalAutoFit(PresentationTextBodyProperties source)
    {
        if (source.NormalAutoFit is not { } normal) return;
        if (source.AutoFitCase != PresentationTextBodyProperties.AutoFitOneofCase.AutoFitMode || source.AutoFitMode != "shrinkText")
            throw Invalid("Presentation normal AutoFit percentages require auto_fit_mode shrinkText.");
        if (normal.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.FontScale1000 && normal.FontScale1000 is < MinFontScale1000 or > MaxFontScale1000)
            throw Invalid("Presentation normal AutoFit font scale must be between 1% and 100%.");
        if (normal.FontScaleCase == PresentationNormalAutoFit.FontScaleOneofCase.NoFontScale && !normal.NoFontScale)
            throw Invalid("Presentation no_font_scale must be true when selected.");
        if (normal.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.LineSpacingReduction1000 && normal.LineSpacingReduction1000 is < MinLineSpacingReduction1000 or > MaxLineSpacingReduction1000)
            throw Invalid("Presentation normal AutoFit line-spacing reduction must be between 0% and 13200%.");
        if (normal.LineSpacingReductionCase == PresentationNormalAutoFit.LineSpacingReductionOneofCase.NoLineSpacingReduction && !normal.NoLineSpacingReduction)
            throw Invalid("Presentation no_line_spacing_reduction must be true when selected.");
    }

    private static string ParseAutoFit(string value) => value switch
    {
        "none" or "shrinkText" or "resizeShape" => value,
        _ => throw Invalid($"Unsupported Presentation AutoFit mode {value}."),
    };

    private static string AnchorName(A.TextAnchoringTypeValues? value) => value is null ? string.Empty :
        value.Value == A.TextAnchoringTypeValues.Top ? "top" :
        value.Value == A.TextAnchoringTypeValues.Center ? "center" :
        value.Value == A.TextAnchoringTypeValues.Bottom ? "bottom" : string.Empty;

    private static A.TextAnchoringTypeValues ParseAnchor(string value) => value switch
    {
        "top" => A.TextAnchoringTypeValues.Top,
        "center" => A.TextAnchoringTypeValues.Center,
        "bottom" => A.TextAnchoringTypeValues.Bottom,
        _ => throw Invalid($"Unsupported Presentation text body anchor {value}."),
    };

    private static string WrapName(A.TextWrappingValues? value) => value is null ? string.Empty :
        value.Value == A.TextWrappingValues.Square ? "square" :
        value.Value == A.TextWrappingValues.None ? "none" : string.Empty;

    private static A.TextWrappingValues ParseWrap(string value) => value switch
    {
        "square" => A.TextWrappingValues.Square,
        "none" => A.TextWrappingValues.None,
        _ => throw Invalid($"Unsupported Presentation text body wrap mode {value}."),
    };

    private static string VerticalTextName(A.TextVerticalValues? value) => value is null ? string.Empty :
        value.Value == A.TextVerticalValues.Horizontal ? "horizontal" :
        value.Value == A.TextVerticalValues.Vertical ? "vertical" :
        value.Value == A.TextVerticalValues.Vertical270 ? "vertical270" : string.Empty;

    private static A.TextVerticalValues ParseVerticalText(string value) => value switch
    {
        "horizontal" => A.TextVerticalValues.Horizontal,
        "vertical" => A.TextVerticalValues.Vertical,
        "vertical270" => A.TextVerticalValues.Vertical270,
        _ => throw Invalid($"Unsupported Presentation vertical text mode {value}."),
    };

    private static string VerticalOverflowName(A.TextVerticalOverflowValues? value) => value is null ? string.Empty :
        value.Value == A.TextVerticalOverflowValues.Overflow ? "overflow" :
        value.Value == A.TextVerticalOverflowValues.Ellipsis ? "ellipsis" :
        value.Value == A.TextVerticalOverflowValues.Clip ? "clip" : string.Empty;

    private static A.TextVerticalOverflowValues ParseVerticalOverflow(string value) => value switch
    {
        "overflow" => A.TextVerticalOverflowValues.Overflow,
        "ellipsis" => A.TextVerticalOverflowValues.Ellipsis,
        "clip" => A.TextVerticalOverflowValues.Clip,
        _ => throw Invalid($"Unsupported Presentation vertical overflow mode {value}."),
    };

    private static string HorizontalOverflowName(A.TextHorizontalOverflowValues? value) => value is null ? string.Empty :
        value.Value == A.TextHorizontalOverflowValues.Overflow ? "overflow" :
        value.Value == A.TextHorizontalOverflowValues.Clip ? "clip" : string.Empty;

    private static A.TextHorizontalOverflowValues ParseHorizontalOverflow(string value) => value switch
    {
        "overflow" => A.TextHorizontalOverflowValues.Overflow,
        "clip" => A.TextHorizontalOverflowValues.Clip,
        _ => throw Invalid($"Unsupported Presentation horizontal overflow mode {value}."),
    };

    private static CodecException Invalid(string message) => new("invalid_presentation_text", message);
    private static CodecException Unsupported(string message) => new("unsupported_presentation_edit", message);
}
