# Complex imported-deck route

Use this reference when an imported PPTX contains dense groups, tables, charts,
embedded media, SVG/EMF pictures, notes, animations, or vendor extensions. It
describes the behavior proved on the six local NASA and SlidesCarnival sample
decks; it is not a promise that those files, or every OOXML feature, is fully
editable.

## Route

```text
import → inspect → classify → resolve one target → edit or stop
       → export to a new path → reimport → review the declared footprint
```

Every visible top-level shape-tree child has a stable source locator and source
revision. Read its classification before editing:

- `typed-editable`: use the typed object operation;
- `native-leaf-editable`: use only an inspect-issued leaf and expected hash;
- `source-derived-reusable`: use the inspected source slide/component reuse
  operation;
- `opaque-preserved`: read its summary and visible text, but do not flatten or
  patch it.

An opaque object may still expose `nativeKind`, bounding box, dependency summary
and read-only `text` through `nativeObject.inspectRecord()`. This is how an
Agent can understand a rich table, vendor shell, or timing-bearing group while
the original XML and relationships remain untouched.

## Safe edits currently demonstrated

- ordinary text, group-child text and bounded geometry through issued native
  leaves;
- embedded pictures and their bounded metadata/placement profiles;
- imported tables with fixed topology, covered cells, style/extension shells,
  and a bounded graphic-frame scale;
- chart title/data and SmartArt text only where their explicit capability is
  present;
- source-bound slide/component reuse and layer ordering when the returned
  capability proves the dependency closure.

Do not add text to an empty or covered imported table cell when no source text
leaf exists. Do not edit a rich multi-run table, relationship-bearing
connector/group, OLE/WPS/animation graph, or irregular SVG through a guessed
selector. Keep it opaque and report the exact blocked capability.

## Review

After every source-bound export, reimport and verify the target value, source
immutability, unchanged non-target parts/relationships, and the declared
SlidePart/media footprint. Render the affected page and at least one unchanged
comparison page when layout can move. A successful import or export alone is
not evidence that a complex deck survived. Windows PowerPoint playback is a
separate host check and must not be implied by portable inspection evidence.
