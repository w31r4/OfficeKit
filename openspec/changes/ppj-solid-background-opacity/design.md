# Design

`background: { "type": "solid", ... }` remains the single PPJ spelling. An
alpha-bearing color or explicit `opacity` lowers to the existing native
`PresentationBackground` plus a presence-aware thousandth-percent field.
Explicit opacity takes precedence over color alpha, matching other PPJ fill
lowering.

The canonical DrawingML profile is intentionally bounded:

```xml
<p:bg>
  <p:bgPr>
    <a:solidFill>
      <a:srgbClr val="0A84FF"><a:alpha val="50000"/></a:srgbClr>
    </a:solidFill>
    <a:effectLst/>
  </p:bgPr>
</p:bg>
```

Only one direct `a:alpha` child is modeled. Other color transforms remain
opaque and fail closed. Style-reference, gradient and image backgrounds keep
their existing independent opacity rules.
