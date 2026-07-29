import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { PPTX_SMARTART_NOTES_COMMENTS_BOUNDARY_FIXTURE } from "./agent-eval-office-fixtures.mjs";

const DIAGRAM_RELATIONSHIP_TYPES = Object.freeze({
  dm: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
  lo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
  qs: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
  cs: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
});

const DIAGRAM_CONTENT_TYPES = Object.freeze({
  "ppt/diagrams/strategy-data.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  "ppt/diagrams/strategy-layout.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  "ppt/diagrams/strategy-style.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  "ppt/diagrams/strategy-colors.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
});

const NOTES_SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const MODERN_COMMENTS_RELATIONSHIP = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
const HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const MODERN_COMMENTS_CONTENT_TYPE = "application/vnd.ms-powerpoint.comments+xml";
const AUTHORS_CONTENT_TYPE = "application/vnd.ms-powerpoint.authors+xml";
const NOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function xmlAttributes(opening = "") {
  const attributes = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    attributes[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return attributes;
}

function drawingTexts(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")));
}

function relationships(xml = "") {
  return [...String(xml).matchAll(/<Relationship\b[^>]*\/?\s*>/g)]
    .map((match) => xmlAttributes(match[0]));
}

function resolvePartPath(sourcePartPath, target = "") {
  if (String(target).startsWith("/")) return String(target).replace(/^\/+/, "");
  return path.posix.normalize(path.posix.join(path.posix.dirname(sourcePartPath), String(target))).replace(/^\/+/, "");
}

function findRelationship(records, id) {
  return records.find((record) => record.Id === id || record.id === id) || null;
}

function relationshipAttribute(record, name) {
  return record?.[name] ?? record?.[name[0].toLowerCase() + name.slice(1)] ?? "";
}

function expectedContentType(overrides, partPath, contentType) {
  return overrides[partPath] === contentType;
}

function parseContentTypeOverrides(xml = "") {
  const overrides = {};
  for (const match of String(xml).matchAll(/<Override\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    const partPath = String(attributes.PartName || attributes.partName || "").replace(/^\/+/, "");
    const contentType = attributes.ContentType || attributes.contentType || "";
    if (partPath) overrides[partPath] = contentType;
  }
  return overrides;
}

function smartArtFrame(slideXml, name) {
  for (const match of String(slideXml).matchAll(/<p:graphicFrame\b[^>]*>[\s\S]*?<\/p:graphicFrame>/g)) {
    const frame = match[0];
    const properties = /<p:cNvPr\b[^>]*\/?\s*>/.exec(frame)?.[0] || "";
    if (xmlAttributes(properties).name !== name) continue;
    const relationIds = /<dgm:relIds\b[^>]*\/?\s*>/.exec(frame)?.[0] || "";
    return { frame, relationIds: xmlAttributes(relationIds) };
  }
  return null;
}

function diagramNodes(dataXml) {
  return [...String(dataXml).matchAll(/<dgm:pt\b[^>]*>[\s\S]*?<\/dgm:pt>/g)].map((match) => ({
    modelId: xmlAttributes(match[0]).modelId || "",
    text: drawingTexts(match[0]).join("\n"),
  }));
}

function modernCommentProfile(commentsXml, authorsXml, fixture) {
  const root = fixture.comment.root;
  const reply = fixture.comment.directReply;
  const rootMatch = [...String(commentsXml).matchAll(/<p188:cm\b([^>]*)>([\s\S]*?)<\/p188:cm>/g)]
    .map((match) => ({ attributes: xmlAttributes(match[1]), body: match[2] }))
    .find((candidate) => candidate.attributes.id === root.id) || null;
  const replyMatch = /<p188:reply\b([^>]*)>([\s\S]*?)<\/p188:reply>/.exec(rootMatch?.body || "");
  const replyAttributes = xmlAttributes(replyMatch?.[1] || "");
  const bodyWithoutReplies = String(rootMatch?.body || "").replace(/<p188:replyLst\b[^>]*>[\s\S]*?<\/p188:replyLst>/g, "");
  const authors = [...String(authorsXml).matchAll(/<p188:author\b[^>]*\/?\s*>/g)]
    .map((match) => xmlAttributes(match[0]));
  const expectedAuthors = [root, reply].every((person) => authors.some((author) => author.id === person.personId
    && author.name === person.author && author.initials === person.initials && author.userId === person.userId));
  return {
    root: Boolean(rootMatch)
      && rootMatch.attributes.authorId === root.personId
      && rootMatch.attributes.status === "active"
      && drawingTexts(bodyWithoutReplies).join("\n").includes(root.text),
    reply: Boolean(replyMatch)
      && replyAttributes.id === reply.id
      && replyAttributes.authorId === reply.personId
      && replyAttributes.status === "active"
      && drawingTexts(replyMatch?.[2] || "").join("\n").includes(reply.text),
    authors: expectedAuthors,
    rootId: rootMatch?.attributes.id || "",
    replyId: replyAttributes.id || "",
    authorCount: authors.length,
  };
}

async function textPart(zip, partPath) {
  const part = zip.file(partPath);
  if (!part) throw new Error(`PPTX SmartArt boundary fixture is missing ${partPath}.`);
  return part.async("text");
}

async function selectedPartHashes(zip, paths) {
  const entries = await Promise.all(paths.map(async (partPath) => {
    const part = zip.file(partPath);
    return [partPath, part ? sha256(await part.async("uint8array")) : null];
  }));
  return Object.fromEntries(entries);
}

/**
 * Reads the PPTX package directly. The public PresentationFile model leaves
 * the connected SmartArt data graph opaque, so this independent evaluator
 * owns proof of the source graph and the unrelated review canaries.
 */
export async function inspectPptxSmartArtNotesCommentsBoundary(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const fixture = PPTX_SMARTART_NOTES_COMMENTS_BOUNDARY_FIXTURE;
  const slideNumber = fixture.smartArt.slideIndex + 1;
  const reviewSlideNumber = fixture.notes.slideIndex + 1;
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const slideRelationshipsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
  const reviewSlidePath = `ppt/slides/slide${reviewSlideNumber}.xml`;
  const reviewRelationshipsPath = `ppt/slides/_rels/slide${reviewSlideNumber}.xml.rels`;
  const notesPath = "ppt/notesSlides/notesSlide1.xml";
  const commentsPath = "ppt/comments/modernComment.xml";
  const authorsPath = "ppt/authors.xml";
  const specialPaths = [
    "[Content_Types].xml",
    slidePath,
    slideRelationshipsPath,
    fixture.smartArt.dataPartPath,
    "ppt/diagrams/strategy-layout.xml",
    "ppt/diagrams/strategy-style.xml",
    "ppt/diagrams/strategy-colors.xml",
    fixture.smartArt.dataRelationshipPath,
    reviewSlidePath,
    reviewRelationshipsPath,
    notesPath,
    commentsPath,
    authorsPath,
  ];
  const [slideXml, slideRelationshipsXml, dataXml, dataRelationshipsXml, reviewSlideXml, reviewRelationshipsXml, notesXml, commentsXml, authorsXml, contentTypesXml] = await Promise.all([
    textPart(zip, slidePath),
    textPart(zip, slideRelationshipsPath),
    textPart(zip, fixture.smartArt.dataPartPath),
    textPart(zip, fixture.smartArt.dataRelationshipPath),
    textPart(zip, reviewSlidePath),
    textPart(zip, reviewRelationshipsPath),
    textPart(zip, notesPath),
    textPart(zip, commentsPath),
    textPart(zip, authorsPath),
    textPart(zip, "[Content_Types].xml"),
  ]);
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths: Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort(),
    partHashes: await selectedPartHashes(zip, specialPaths),
    slideXml,
    slideRelationships: relationships(slideRelationshipsXml),
    dataXml,
    dataRelationships: relationships(dataRelationshipsXml),
    reviewSlideXml,
    reviewSlideRelationships: relationships(reviewRelationshipsXml),
    notesXml,
    commentsXml,
    authorsXml,
    contentTypeOverrides: parseContentTypeOverrides(contentTypesXml),
  };
}

export function pptxSmartArtNotesCommentsProfile(source) {
  const fixture = PPTX_SMARTART_NOTES_COMMENTS_BOUNDARY_FIXTURE;
  const smartArt = smartArtFrame(source.slideXml, fixture.smartArt.name);
  const expectedDiagramEdges = Object.entries(DIAGRAM_RELATIONSHIP_TYPES).map(([attribute, type]) => {
    const relationship = findRelationship(source.slideRelationships, smartArt?.relationIds?.[attribute]);
    return {
      attribute,
      id: smartArt?.relationIds?.[attribute] || "",
      type: relationshipAttribute(relationship, "Type"),
      target: resolvePartPath(`ppt/slides/slide${fixture.smartArt.slideIndex + 1}.xml`, relationshipAttribute(relationship, "Target")),
      valid: relationshipAttribute(relationship, "Type") === type,
    };
  });
  const expectedDiagramParts = Object.keys(DIAGRAM_CONTENT_TYPES);
  const expectedTargets = {
    dm: fixture.smartArt.dataPartPath,
    lo: "ppt/diagrams/strategy-layout.xml",
    qs: "ppt/diagrams/strategy-style.xml",
    cs: "ppt/diagrams/strategy-colors.xml",
  };
  const diagramEdges = expectedDiagramEdges.map((edge) => ({
    ...edge,
    valid: edge.valid && edge.target === expectedTargets[edge.attribute],
  }));
  const externalLinks = source.dataRelationships.filter((relationship) => relationshipAttribute(relationship, "Type") === HYPERLINK_RELATIONSHIP);
  const externalLink = externalLinks[0] || null;
  const externalDataRelationship = externalLinks.length === 1
    && relationshipAttribute(externalLink, "TargetMode") === "External"
    && relationshipAttribute(externalLink, "Target") === fixture.smartArt.externalTarget;
  const nodes = diagramNodes(source.dataXml);
  const expectedNodes = fixture.smartArt.nodes.every((expected) => nodes.some((node) => node.modelId === expected.id && node.text === expected.text));
  const reviewNotes = drawingTexts(source.notesXml).join("\n").includes(fixture.notes.text);
  const notesRelationships = source.reviewSlideRelationships.filter((relationship) => relationshipAttribute(relationship, "Type") === NOTES_SLIDE_RELATIONSHIP);
  const commentsRelationships = source.reviewSlideRelationships.filter((relationship) => relationshipAttribute(relationship, "Type") === MODERN_COMMENTS_RELATIONSHIP);
  const reviewGraph = notesRelationships.length === 1
    && resolvePartPath(`ppt/slides/slide${fixture.notes.slideIndex + 1}.xml`, relationshipAttribute(notesRelationships[0], "Target")) === "ppt/notesSlides/notesSlide1.xml"
    && commentsRelationships.length === 1
    && resolvePartPath(`ppt/slides/slide${fixture.notes.slideIndex + 1}.xml`, relationshipAttribute(commentsRelationships[0], "Target")) === "ppt/comments/modernComment.xml";
  const comments = modernCommentProfile(source.commentsXml, source.authorsXml, fixture);
  const contentTypes = Object.entries(DIAGRAM_CONTENT_TYPES).every(([partPath, type]) => expectedContentType(source.contentTypeOverrides, partPath, type))
    && expectedContentType(source.contentTypeOverrides, "ppt/notesSlides/notesSlide1.xml", NOTES_CONTENT_TYPE)
    && expectedContentType(source.contentTypeOverrides, "ppt/comments/modernComment.xml", MODERN_COMMENTS_CONTENT_TYPE)
    && expectedContentType(source.contentTypeOverrides, "ppt/authors.xml", AUTHORS_CONTENT_TYPE);
  const slidePaths = source.paths.filter((partPath) => /^ppt\/slides\/slide\d+\.xml$/i.test(partPath));
  const requiredParts = expectedDiagramParts.concat([
    fixture.smartArt.dataRelationshipPath,
    "ppt/notesSlides/notesSlide1.xml",
    "ppt/comments/modernComment.xml",
    "ppt/authors.xml",
  ]).every((partPath) => source.paths.includes(partPath));
  return {
    ok: slidePaths.length === fixture.slides.length
      && Boolean(smartArt)
      && diagramEdges.length === 4
      && diagramEdges.every((edge) => edge.valid)
      && expectedNodes
      && externalDataRelationship
      && reviewGraph
      && reviewNotes
      && comments.root
      && comments.reply
      && comments.authors
      && contentTypes
      && requiredParts,
    slideCount: slidePaths.length,
    smartArt: {
      name: fixture.smartArt.name,
      found: Boolean(smartArt),
      edges: diagramEdges,
      nodeCount: nodes.length,
      nodes,
      externalDataRelationship,
      externalTarget: relationshipAttribute(externalLink, "Target"),
    },
    review: {
      graph: reviewGraph,
      notes: reviewNotes,
      modernComments: comments,
    },
    contentTypes,
    requiredParts,
  };
}
