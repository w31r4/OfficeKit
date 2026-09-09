using System.Collections;
using System.Security.Cryptography;
using System.Text;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Copies the existing native IR without reducing it to painter-supported fields.
// Export authorization and opaque XML are intentionally not preview transport.
internal static class PpjPreviewSceneBuilder
{
    internal const uint Version = 1;
    internal const int MaxDepth = 128;

    internal static PresentationPreviewScene Build(
        PresentationArtifact presentation,
        PresentationPreviewSceneOrigin origin,
        string programSha256,
        string candidateSha256,
        IEnumerable<PresentationPreviewNodeBinding> bindings,
        IEnumerable<Asset> assets,
        EffectiveCodecLimits limits)
    {
        var budget = NewBudget(limits);
        return Seal((PresentationArtifact)Copy(presentation, budget, 0), origin,
            programSha256, candidateSha256, bindings, assets, budget);
    }

    // Instantiated only by opt-in callers. A slide is copied exactly once when
    // the writer materializes it; sealing does not clone that graph again.
    internal sealed class Collector
    {
        private readonly PresentationArtifact snapshot;
        private readonly CopyBudget budget;
        private int nextSlide;
        private bool completed;

        internal Collector(PresentationArtifact headers, EffectiveCodecLimits limits)
        {
            if (headers.Slides.Any(slide => slide.Elements.Count != 0))
                throw Invalid("Authored scene collection requires unmaterialized slide headers.");
            budget = NewBudget(limits);
            snapshot = (PresentationArtifact)Copy(headers, budget, 0);
        }

        internal void CaptureSlide(int index, PresentationSlide slide)
        {
            if (completed || index != nextSlide || index >= snapshot.Slides.Count || snapshot.Slides[index].Id != slide.Id)
                throw Invalid("Scene slides must be captured once in writer order with matching page identity.");
            snapshot.Slides[index] = (PresentationSlide)Copy(slide, budget, 0);
            nextSlide++;
        }

        internal PresentationPreviewScene Complete(string programSha256, string candidateSha256,
            IEnumerable<PresentationPreviewNodeBinding> bindings, IEnumerable<Asset> assets)
        {
            if (completed || nextSlide != snapshot.Slides.Count)
                throw Invalid("Scene collection is incomplete or has already been sealed.");
            // Even failed sealing is terminal: partially added bindings/assets
            // must never be reused as the start of another complete receipt.
            completed = true;
            return Seal(snapshot, PresentationPreviewSceneOrigin.AuthoredLowering,
                programSha256, candidateSha256, bindings, assets, budget);
        }
    }

    private static CopyBudget NewBudget(EffectiveCodecLimits limits) =>
        new(Math.Min(limits.MaxUncompressedBytes, (ulong)CodecWireProtocol.AbsoluteRequestLimit));

    private static PresentationPreviewScene Seal(PresentationArtifact snapshot,
        PresentationPreviewSceneOrigin origin, string programSha256, string candidateSha256,
        IEnumerable<PresentationPreviewNodeBinding> bindings, IEnumerable<Asset> assets, CopyBudget budget)
    {
        if (origin is not (PresentationPreviewSceneOrigin.AuthoredLowering or PresentationPreviewSceneOrigin.CandidateImport))
            throw Invalid("Scene origin must identify authored lowering or the actual candidate import.");
        RequireHash(programSha256);
        RequireHash(candidateSha256);
        var scene = new PresentationPreviewScene
        {
            Version = Version,
            Origin = origin,
            ProgramSha256 = programSha256,
            CandidateSha256 = candidateSha256,
            Presentation = snapshot,
        };
        var paths = new HashSet<string>(StringComparer.Ordinal);
        foreach (var binding in bindings)
        {
            if (string.IsNullOrEmpty(binding.PageId) || string.IsNullOrEmpty(binding.NativeId) ||
                !binding.ScenePath.StartsWith("$.presentation.", StringComparison.Ordinal) ||
                !binding.ProgramPath.StartsWith("$", StringComparison.Ordinal) ||
                !paths.Add(binding.ScenePath) ||
                binding.Attribution is not (PresentationPreviewAttribution.Direct or PresentationPreviewAttribution.Generated or PresentationPreviewAttribution.Unmapped) ||
                (binding.Attribution != PresentationPreviewAttribution.Unmapped && string.IsNullOrEmpty(binding.SemanticId)))
                throw Invalid("Scene bindings require unique scene paths and explicit semantic attribution.");
            scene.Bindings.Add((PresentationPreviewNodeBinding)Copy(binding, budget, 0));
        }
        // Traversal order is meaningful for nodes; asset input order is not.
        var assetIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var asset in assets.OrderBy(asset => asset.Id, StringComparer.Ordinal))
        {
            if (string.IsNullOrEmpty(asset.Id) || string.IsNullOrEmpty(asset.ContentType) || !assetIds.Add(asset.Id))
                throw Invalid("Scene asset identities must be non-empty and unique.");
            // Asset ingress accepts either hex case. Normalize evidence without
            // changing the caller's receipt or rejecting an otherwise valid build.
            var assetHash = asset.Sha256.ToLowerInvariant();
            RequireHash(assetHash);
            if (asset.Data.IsEmpty || Hash(asset.Data.Span) != assetHash)
                throw new CodecException("preview_scene_asset_mismatch", "Scene asset bytes do not match their declared identity.", asset.Id);
            var reference = new PresentationPreviewAssetReference
            {
                NativeId = asset.Id, ContentType = asset.ContentType, Sha256 = assetHash,
            };
            scene.Assets.Add((PresentationPreviewAssetReference)Copy(reference, budget, 0));
        }
        // Include envelope framing and the final digest in the exact wire cap.
        scene.Sha256 = new string('0', 64);
        if ((ulong)scene.CalculateSize() > budget.MaxBytes) throw BudgetExceeded("Serialized scene exceeds its byte budget.");
        scene.Sha256 = string.Empty;
        scene.Sha256 = Hash(Serialize(scene));
        return scene;
    }

    internal static byte[] Serialize(PresentationPreviewScene scene)
    {
        var bytes = new byte[scene.CalculateSize()];
        using var output = new CodedOutputStream(bytes) { Deterministic = true };
        scene.WriteTo(output);
        output.CheckNoSpaceLeft();
        return bytes;
    }

    private static IMessage Copy(IMessage source, CopyBudget budget, int depth)
    {
        if (depth > MaxDepth) throw BudgetExceeded("Scene exceeds the nested-message depth budget.");
        budget.AddBytes(8);
        if (source is PresentationElement or PresentationPlaceholder) budget.AddNode();
        var copy = source.Descriptor.Parser.ParseFrom(ByteString.Empty);
        foreach (var field in source.Descriptor.Fields.InFieldNumberOrder())
        {
            if (Excluded(field)) continue;
            if (field.IsMap) throw Invalid($"Unclassified scene map field: {field.FullName}.");
            if (field.HasPresence && !field.Accessor.HasValue(source)) continue;
            var value = field.Accessor.GetValue(source);
            if (field.IsRepeated)
            {
                var target = (IList)field.Accessor.GetValue(copy);
                foreach (var item in (IEnumerable)value)
                    target.Add(CopyValue(item, budget, depth));
            }
            else if (value is not null)
                field.Accessor.SetValue(copy, CopyValue(value, budget, depth));
        }
        return copy;
    }

    private static object CopyValue(object value, CopyBudget budget, int depth)
    {
        if (value is IMessage message) return Copy(message, budget, depth + 1);
        // There are no inline bytes in current Presentation IR. A future such
        // field needs an explicit asset/opaque policy, not silent omission.
        if (value is ByteString) throw Invalid("Unclassified inline binary field in preview scene.");
        budget.AddBytes(value is string text ? checked((ulong)Encoding.UTF8.GetByteCount(text) + 8) : 16UL);
        return value;
    }

    private static bool Excluded(FieldDescriptor field)
    {
        if (field.FieldType == FieldType.Message &&
            (field.MessageType.Name.EndsWith("SourceBinding", StringComparison.Ordinal) ||
             field.MessageType.Name.EndsWith("Capability", StringComparison.Ordinal) ||
             field.MessageType.Name == nameof(PresentationElementDeletion))) return true;
        return (field.ContainingType.Name == nameof(PresentationOpaqueElement) && field.Name == "raw_xml") ||
            (field.ContainingType.Name is nameof(PresentationOleWorkbook) or nameof(PresentationOleOfficePackage) &&
                field.Name == "replacement_asset_id");
    }

    private static void RequireHash(string value)
    {
        if (value.Length != 64 || value.Any(c => c is not (>= '0' and <= '9') and not (>= 'a' and <= 'f')))
            throw Invalid("Scene identities require lowercase SHA-256 digests.");
    }

    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_preview_scene", message);
    private static CodecException BudgetExceeded(string message) => new("preview_scene_budget_exceeded", message);

    private sealed class CopyBudget(ulong maxBytes)
    {
        internal ulong MaxBytes { get; } = maxBytes;
        private ulong retainedBytes;
        private int nodes;

        internal void AddBytes(ulong count)
        {
            // Conservative in-memory copy budget, separate from the exact wire
            // size check. Count before allocation and never return truncation.
            if (count > MaxBytes || retainedBytes > MaxBytes - count)
                throw BudgetExceeded("Scene exceeds its retained-value byte budget.");
            retainedBytes += count;
        }

        internal void AddNode()
        {
            if (++nodes > PpjProgramValidator.MaxExpandedElements)
                throw BudgetExceeded("Scene exceeds the expanded visual-node budget.");
        }
    }
}
