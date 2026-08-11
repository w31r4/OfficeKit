import { ndjson } from "../shared/inspection.mjs";

const GENERIC_LINK_TEXT = new Set([
  "click here",
  "here",
  "learn more",
  "link",
  "more",
  "read more",
  "点击这里",
  "更多",
  "链接",
]);

function locator(block, index) {
  return {
    blockIndex: index,
    id: block.id,
    blockKind: block.kind,
    ...(block.name ? { name: block.name } : {}),
  };
}

function issue(block, index, type, message, details = {}) {
  return {
    kind: "accessibilityIssue",
    artifactKind: "document",
    type,
    severity: "error",
    message,
    ...locator(block, index),
    ...details,
  };
}

function manualCheck(block, index, type, message, details = {}) {
  return {
    kind: "accessibilityManualCheck",
    artifactKind: "document",
    type,
    severity: "manual",
    message,
    ...locator(block, index),
    ...details,
  };
}

function namedHeadingLevel(style = {}) {
  for (const value of [style.id, style.name]) {
    const match = /^heading[\s_-]*([1-9])$/iu.exec(String(value || "").trim());
    if (match) return Number(match[1]);
  }
  return undefined;
}

function headingLevel(document, block) {
  if (Object.hasOwn(block.paragraphFormat || {}, "outlineLevel")) {
    const direct = block.paragraphFormat.outlineLevel;
    return Number.isInteger(direct) && direct >= 0 && direct <= 8 ? direct + 1 : undefined;
  }
  const effective = document.styles?.effective?.(block.styleId) || {};
  if (Number.isInteger(effective.outlineLevel) && effective.outlineLevel >= 0 && effective.outlineLevel <= 8) {
    return effective.outlineLevel + 1;
  }
  return namedHeadingLevel(effective) ?? namedHeadingLevel({ id: block.styleId });
}

function hyperlinkNeedsPurposeReview(text) {
  const normalized = String(text || "").trim().toLocaleLowerCase("en-US");
  return GENERIC_LINK_TEXT.has(normalized) || /^(?:https?:\/\/|www\.)\S+$/iu.test(normalized);
}

export function auditDocumentAccessibility(document, options = {}) {
  if (!document || typeof document !== "object" || !Array.isArray(document.blocks)) {
    throw new TypeError("Document accessibility audit requires a DocumentModel-like object.");
  }
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    throw new TypeError("Document accessibility audit options must be an object.");
  }

  const issues = [];
  const manualChecks = [];
  const summary = {
    blocks: document.blocks.length,
    headings: 0,
    headingLevelSkips: 0,
    images: 0,
    imagesMissingAltText: 0,
    tables: 0,
    tablesWithoutHeaderRows: 0,
    tablesWithoutAccessibilityMetadata: 0,
    links: 0,
    linksRequiringPurposeReview: 0,
  };
  let previousHeadingLevel = 0;

  document.blocks.forEach((block, index) => {
    if (!block || typeof block !== "object") {
      throw new TypeError("Document accessibility audit blocks must be objects.");
    }

    if (block.kind === "paragraph") {
      const level = headingLevel(document, block);
      if (level === undefined) return;
      summary.headings += 1;
      if (level > previousHeadingLevel + 1) {
        summary.headingLevelSkips += 1;
        issues.push(issue(
          block,
          index,
          "headingLevelSkipped",
          `Heading ${block.name || block.id} jumps from level ${previousHeadingLevel || "body"} to level ${level}.`,
          { headingLevel: level, previousHeadingLevel },
        ));
      }
      previousHeadingLevel = level;
      return;
    }

    if (block.kind === "image") {
      summary.images += 1;
      if (!String(block.alt || "").trim()) {
        summary.imagesMissingAltText += 1;
        issues.push(issue(
          block,
          index,
          "imageAltTextMissing",
          `Image ${block.name || block.id} has no non-empty alternative text.`,
        ));
      }
      return;
    }

    if (block.kind === "table") {
      summary.tables += 1;
      if (!Number.isInteger(block.headerRowCount) || block.headerRowCount < 1) {
        summary.tablesWithoutHeaderRows += 1;
        issues.push(issue(
          block,
          index,
          "tableHeaderRowMissing",
          `Table ${block.name || block.id} has no declared repeating header-row prefix.`,
          { rows: block.rows, columns: block.columns },
        ));
      }
      if (!block.accessibility?.title && !block.accessibility?.description) {
        summary.tablesWithoutAccessibilityMetadata += 1;
        manualChecks.push(manualCheck(
          block,
          index,
          "tablePurposeAndDescription",
          `Table ${block.name || block.id} has no non-visible title or description; review whether assistive users need one and whether the declared header rows express the real table semantics.`,
          { headerRowCount: block.headerRowCount },
        ));
      }
      return;
    }

    if (block.kind !== "hyperlink") return;
    summary.links += 1;
    const text = String(block.text || "").trim();
    if (!text) {
      issues.push(issue(
        block,
        index,
        "hyperlinkTextMissing",
        `Hyperlink ${block.name || block.id} has no visible text.`,
      ));
      return;
    }
    if (!hyperlinkNeedsPurposeReview(text)) return;
    summary.linksRequiringPurposeReview += 1;
    manualChecks.push(manualCheck(
      block,
      index,
      "hyperlinkPurpose",
      `Hyperlink ${block.name || block.id} uses generic or raw-URL text; review its purpose in surrounding context.`,
      { text, target: block.anchor ? `#${block.anchor}` : block.url },
    ));
  });

  const records = [...issues, ...manualChecks];
  return {
    kind: "documentAccessibilityAudit",
    artifactKind: "document",
    machineCheckPassed: issues.length === 0,
    conformanceClaimed: false,
    manualReviewRequired: manualChecks.length > 0,
    summary,
    issues,
    manualChecks,
    ...ndjson(records, options.maxChars ?? Infinity),
  };
}
