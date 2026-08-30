# Design

## Language token

The public value uses the existing PPJ `languageTag` profile: 2–63 characters,
an alphabetic primary subtag, and bounded alphanumeric subtags. OfficeKit
preserves caller spelling and does not claim full IANA registry validation.

```json
{
  "text": {
    "paragraphs": [{
      "runs": [
        { "text": "关键结论", "style": { "language": "zh-CN" } },
        { "text": " / Evidence", "style": { "language": "en-US" } }
      ]
    }]
  }
}
```

## Native mapping

`PresentationTextRun.language` maps to direct `a:rPr/@lang`.
`PresentationTextStyle.language` maps to direct `a:defRPr/@lang`. Newly authored
runs retain the existing `en-US` fallback only when PPJ does not specify a
language; imported runs preserve their explicit value.

## Source-bound edit

A canonical explicit `a:rPr/@lang` on one editable text run issues
`fontLanguage`. The Edit Plan binds source, slide, element, text leaf and run
attribute hashes, then token-splices only the escaped `lang` value. Missing,
inherited, duplicate or malformed language state does not issue a capability.

## Verification

Extend the existing integrated PPJ contract rather than adding a new fixture or
language matrix.
