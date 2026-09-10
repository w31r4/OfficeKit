// Independent fixed 240x120 DrawingML references. No production geometry or
// default-table import; actual presets are compared with literal custom paths.
export const roundedCornerPresets = ["round1Rect", "round2DiagRect", "round2SameRect", "snipRoundRect"];
export function roundedCornerReference(preset, values = []) {
  const defaults = preset === "round1Rect" ? [16667] : preset === "snipRoundRect" ? [16667,16667] : [16667,0];
  const radii = { 0:0, 16667:20.0004, 25000:30, 50000:60 };
  const [a,b] = defaults.map((value,i) => radii[values[i] ?? value]);
  if (![a,...(b === undefined ? [] : [b])].every(Number.isFinite)) throw new Error("Missing independent rounded-corner fixture");
  const inset = { 0:0, 20.0004:5.857917156, 30:8.7867, 60:17.5734 };
  const ia = inset[a], ib = inset[b], im = Math.max(ia,ib ?? ia);
  const m = (x,y) => ({op:"moveTo",x,y}), l = (x,y) => ({op:"lineTo",x,y});
  const arc = (r,startAngle) => r ? [{op:"arcTo",radiusX:r,radiusY:r,startAngle,sweepAngle:90}] : [];
  let commands, edges;
  if (preset === "round1Rect") {
    commands=[m(0,0),l(240-a,0),...arc(a,270),l(240,120),l(0,120)];
    edges=[0,0,240-ia,120];
  } else if (preset === "round2DiagRect") {
    commands=[m(a,0),l(240-b,0),...arc(b,270),l(240,120-a),...arc(a,0),l(b,120),...arc(b,90),l(0,a),...arc(a,180)];
    edges=[im,im,240-im,120-im];
  } else if (preset === "round2SameRect") {
    commands=[m(a,0),l(240-a,0),...arc(a,270),l(240,120-b),...arc(b,0),l(b,120),...arc(b,90),l(0,a),...arc(a,180)];
    edges=[im,ia,240-im,120-ib];
  } else if (preset === "snipRoundRect") {
    commands=[m(a,0),l(240-b,0),l(240,b),l(240,120),l(0,120),l(0,a),...arc(a,180)];
    edges=[ia,ia,240-b/2,120];
  } else throw new Error(`Unknown rounded-corner fixture ${preset}`);
  // The native custom-path writer stores 1/1000 of a declared viewport unit.
  // Enlarge the reference viewport (not the physical frame) so the default
  // 20.0004pt radius survives compilation instead of being rounded to 20pt.
  const scale = 100;
  commands = commands.map(command => Object.fromEntries(Object.entries(command).map(([key,value]) =>
    [key,["x","y","radiusX","radiusY"].includes(key) ? value * scale : value])));
  return {commands:[...commands,{op:"close"}],edges,viewBox:{x:0,y:0,width:240*scale,height:120*scale}};
}

// The contour fixture contains one 16pt black F0 run, default physical insets
// 7.2/3.6pt and a fixed (100,100) frame. Emit its independent literal SVG
// reference directly: a native custom text rectangle would round the exact
// preset inset to whole EMUs, changing the reflected group's raster bounds.
// Do not copy positions from actual output or round the actual preset geometry.
export function roundedCornerGlyphReference(preset, values = []) {
  const {edges:[left,top,right]} = roundedCornerReference(preset,values);
  const text = `<text x="${100+left+7.2}" y="${100+top+3.6+16}" text-anchor="start" xml:space="preserve"><tspan text-decoration="none" dy="0" font-family="sans-serif" font-size="16" font-weight="normal" font-style="normal" fill="#000000" fill-opacity="1">F0</tspan></text>`;
  return ["round2DiagRect","snipRoundRect"].includes(preset)
    ? `<g transform="translate(${200+left+right} 0) scale(-1 1)">${text}</g>` : text;
}
