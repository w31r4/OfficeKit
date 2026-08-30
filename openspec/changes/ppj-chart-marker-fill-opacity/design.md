# Design

The marker keeps its existing direct RGB fill plus a presence-aware opacity
field. This avoids a second general chart-paint union for a marker profile that
only owns direct solid color.

Canonical native form:

```xml
<c:marker>
  <c:spPr>
    <a:solidFill>
      <a:srgbClr val="0A84FF"><a:alpha val="50000"/></a:srgbClr>
    </a:solidFill>
  </c:spPr>
</c:marker>
```

Other color transforms, gradient/image marker paint and picture markers remain
source-owned. Projection emits the compact alpha-bearing PPJ color because the
marker schema has only one fill spelling.
