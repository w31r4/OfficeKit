// Fixed 240x120 references from the pinned DrawingML definitions. Deliberately
// no production geometry imports: native preset paint is compared with a
// separately compiled literal path and independently specified text rectangle.
export const cutCornerPresets = ["octagon", "snip1Rect", "snip2DiagRect", "snip2SameRect"];
export const adjustmentList = value => value === undefined ? [] : Array.isArray(value) ? value : [value];
export function cutCornerReference(preset, adjustment) {
  const key = adjustmentList(adjustment).join(",");
  let cuts, edges;
  if (preset === "octagon") {
    const d = { "": 35.1468, "25000": 30, "50000": 60, "0": 0 }[key];
    cuts = [d,d,d,d]; edges = [d/2,d/2,240-d/2,120-d/2];
  } else if (preset === "snip1Rect") {
    const d = { "": 20.0004, "25000": 30, "50000": 60, "0": 0 }[key];
    cuts = [0,d,0,0]; edges = [0,d/2,240-d/2,120];
  } else if (preset === "snip2DiagRect") {
    [cuts,edges] = {
      "": [[0,20.0004,0,20.0004],[10.0002,10.0002,229.9998,109.9998]],
      "25000,50000": [[30,60,30,60],[30,30,210,90]],
      "50000,25000": [[60,30,60,30],[30,30,210,90]],
      "25000,16667": [[30,20.0004,30,20.0004],[15,15,225,105]],
      "0,0": [[0,0,0,0],[0,0,240,120]],
    }[key] ?? [];
  } else if (preset === "snip2SameRect") {
    [cuts,edges] = {
      "": [[20.0004,20.0004,0,0],[10.0002,10.0002,229.9998,120]],
      "25000,50000": [[30,30,60,60],[30,15,210,90]],
      "50000,25000": [[60,60,30,30],[30,30,210,105]],
      "25000,0": [[30,30,0,0],[15,15,225,120]],
      "0,0": [[0,0,0,0],[0,0,240,120]],
    }[key] ?? [];
  }
  if (!cuts?.every(Number.isFinite) || !edges?.every(Number.isFinite))
    throw new Error(`Missing independent cut-corner reference: ${preset}/${key}`);
  const [tl,tr,br,bl] = cuts;
  const vertices = [[tl,0],[240-tr,0],[240,tr],[240,120-br],[240-br,120],[bl,120],[0,120-bl],[0,tl]];
  // The octagon starts on its left edge in the definition, not its top edge.
  if (preset === "octagon") vertices.unshift(vertices.pop());
  const points = vertices.filter((p,i) => i === 0 || p.some((v,a) => v !== vertices[i-1][a]));
  if (points.at(-1).every((v,a) => v === points[0][a])) points.pop();
  return { points, edges };
}
