import { attrEscape, xmlEscape } from "../shared/xml.mjs";

export function renderWorksheetImageSvg(image) {
  const position = image.position;
  const filters = [];
  if (image.effects?.grayscale === true) filters.push("grayscale(1)");
  const brightness = Number(image.effects?.brightnessPercent);
  const contrast = Number(image.effects?.contrastPercent);
  const opacity = Number(image.effects?.opacityPercent);
  if (Number.isFinite(brightness) && brightness >= -100 && brightness <= 100) filters.push(`brightness(${1 + brightness / 100})`);
  if (Number.isFinite(contrast) && contrast >= -100 && contrast <= 100) filters.push(`contrast(${1 + contrast / 100})`);
  const degrees = Number(image.transform?.rotationDegrees);
  const flipHorizontal = image.transform?.flipHorizontal === true ? -1 : 1;
  const flipVertical = image.transform?.flipVertical === true ? -1 : 1;
  const centerX = position.left + position.width / 2;
  const centerY = position.top + position.height / 2;
  const transform = (Number.isFinite(degrees) || flipHorizontal < 0 || flipVertical < 0)
    ? ` transform="translate(${centerX} ${centerY}) rotate(${Number.isFinite(degrees) ? degrees : 0}) scale(${flipHorizontal} ${flipVertical}) translate(${-centerX} ${-centerY})"`
    : "";
  const visual = `${filters.length ? ` style="filter:${attrEscape(filters.join(" "))}"` : ""}${Number.isFinite(opacity) && opacity >= 0 && opacity <= 100 ? ` opacity="${opacity / 100}"` : ""}${transform}`;
  const decorative = image.accessibility?.decorative === true;
  const semantics = decorative
    ? ""
    : `${image.accessibility?.title ? `<title>${xmlEscape(image.accessibility.title)}</title>` : ""}${image.accessibility?.description ? `<desc>${xmlEscape(image.accessibility.description)}</desc>` : ""}`;
  const groupAttributes = decorative ? " aria-hidden=\"true\"" : image.accessibility ? " role=\"img\"" : "";
  const content = image.dataUrl
    ? `<image href="${attrEscape(image.dataUrl)}" x="${position.left}" y="${position.top}" width="${position.width}" height="${position.height}" preserveAspectRatio="xMidYMid meet"${visual}/>`
    : `<rect x="${position.left}" y="${position.top}" width="${position.width}" height="${position.height}" fill="#fef3c7" stroke="#f59e0b"${visual}/><text x="${position.left + 8}" y="${position.top + 20}" font-family="Arial" font-size="12" fill="#92400e">${xmlEscape(image.prompt || image.name)}</text>`;
  return `<g${groupAttributes}>${semantics}${content}</g>`;
}
