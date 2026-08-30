# Design

## PPJ surface

```json
{
  "type": "text",
  "text": {
    "paragraphs": [{
      "runs": [{
        "text": "Decision threshold",
        "style": { "highlight": "#FFF2CC" }
      }]
    }]
  }
}
```

`highlight` participates in the existing text-style cascade: named default,
inline text-box default, then inline run style. It accepts PPJ direct colors or
declared theme tokens. Highlight alpha is rejected because canonical DrawingML
highlight does not share the ordinary text-color alpha contract.

## Projection boundary

Direct RGB `a:highlight` becomes typed PPJ. An imported theme highlight remains
visible as a capability-issued `fontHighlightScheme` leaf because imported PPJ
does not yet declare a complete source theme-token catalog. Effect-bearing,
transformed or malformed highlight graphs remain source owned.

## Source-bound lowering

This change does not widen ordinary text replacement into arbitrary style
editing. An Agent changes an imported highlight only through an issued native
leaf with the expected source and leaf hashes; the native writer token-splices
the exact color child and re-proves the source graph.

## Verification

Extend the existing comprehensive PPJ build/import contract. It proves authored
native output and direct-RGB projection without adding a fixture or matrix.
