import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, Presentation, PresentationFile, renderArtifact } from "office-kit";
import { playwrightRenderer } from "office-kit/renderers/playwright";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredPath(value, name) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${name} must be a non-empty path.`);
  return path.resolve(value);
}

function chartByName(presentation, name) {
  const matches = presentation.slides.getItem(0).charts.items.filter((chart) => chart.name === name);
  if (matches.length !== 1) throw new Error(`Expected exactly one chart named ${JSON.stringify(name)}; found ${matches.length}.`);
  return matches[0];
}

async function trendlineInventory(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((entry) => /(?:^|\/)charts\/chart\d+\.xml$/i.test(entry)).sort();
  return Promise.all(paths.map(async (partPath) => {
    const xml = await zip.file(partPath).async("text");
    return {
      partPath,
      title: /<c:title>[\s\S]*?<a:t>([^<]+)<\/a:t>/.exec(xml)?.[1] || "",
      types: [...xml.matchAll(/<c:trendlineType val="([^"]+)"\s*\/>/g)].map((match) => match[1]),
      names: [...xml.matchAll(/<c:name>([^<]+)<\/c:name>/g)].map((match) => match[1]),
      sha256: sha256(Buffer.from(xml)),
    };
  }));
}

function createTrendlineDeck() {
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const slide = presentation.slides.add({ name: "Trendline evidence" });
  slide.charts.add("line", {
    name: "pipeline-trend",
    title: "Pipeline trajectory",
    position: { left: 45, top: 85, width: 570, height: 545 },
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [{
      name: "Pipeline",
      values: [42, 51, 63, 78],
      line: { fill: "#2563EB", width: 2 },
      marker: { symbol: "circle", size: 7, fill: "#2563EB" },
      trendlines: [
        {
          type: "linear",
          name: "Pipeline projection",
          forward: 0.5,
          displayEquation: true,
          displayRSquared: true,
          line: { fill: "#7C3AED", width: 1.5, style: "dash" },
        },
        { type: "movingAvg", name: "Two-quarter average", period: 2 },
        { type: "poly", name: "Pipeline curve", order: 2 },
      ],
    }],
    xAxis: { title: "Quarter" },
    yAxis: { title: "Pipeline ($M)", min: 0, max: 100, majorUnit: 20 },
    legend: false,
  });
  slide.charts.add("combo", {
    name: "revenue-margin-trend",
    title: "Revenue and margin",
    position: { left: 665, top: 85, width: 570, height: 545 },
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "Revenue", chartType: "bar", values: [35, 43, 54, 68], color: "#0EA5E9" },
      {
        name: "Margin",
        chartType: "line",
        axisGroup: "secondary",
        values: [12, 15, 18, 22],
        line: { fill: "#16A34A", width: 2 },
        marker: { symbol: "diamond", size: 7, fill: "#16A34A" },
        trendlines: [{ type: "exp", name: "Margin projection", forward: 0.5, line: { fill: "#F97316", width: 1.5, style: "dot" } }],
      },
    ],
    axes: {
      category: { title: "Quarter" },
      value: { title: "Revenue ($M)" },
      secondary: { category: { title: "Quarter" }, value: { title: "Margin (%)", min: 0, max: 40, majorUnit: 10 } },
    },
    legend: true,
  });
  return presentation;
}

export async function createAndEditTrendlineDeck({ outputPath, previewPath, auditPath }) {
  const output = requiredPath(outputPath, "outputPath");
  const preview = requiredPath(previewPath, "previewPath");
  const auditPathname = requiredPath(auditPath, "auditPath");
  if (new Set([output, preview, auditPathname]).size !== 3) throw new Error("outputPath, previewPath, and auditPath must be distinct.");
  const temporary = [output, preview, auditPathname].map((entry) => `${entry}.tmp-${process.pid}-${Date.now()}`);
  await Promise.all([output, preview, auditPathname].map((entry) => fs.mkdir(path.dirname(entry), { recursive: true })));
  try {
    const authored = createTrendlineDeck();
    const authoredVerification = authored.verify({ visualQa: true });
    assert.equal(authoredVerification.ok, true, authoredVerification.ndjson);
    const first = await PresentationFile.exportPptx(authored);
    const firstBytes = new Uint8Array(await first.arrayBuffer());
    const firstInventory = await trendlineInventory(firstBytes);
    assert.deepEqual(firstInventory.map((entry) => entry.types), [["linear", "movingAvg", "poly"], ["exp"]]);

    const imported = await PresentationFile.importPptx(new FileBlob(firstBytes, { type: PPTX_MIME, name: "trendline-source.pptx" }));
    const pipeline = chartByName(imported, "pipeline-trend");
    const margin = chartByName(imported, "revenue-margin-trend");
    pipeline.series[0].trendlines[0].name = "Updated pipeline projection";
    pipeline.series[0].trendlines[0].forward = 1.5;
    pipeline.series[0].trendlines[0].line.fill = "#0EA5E9";
    margin.series[1].trendlines[0].name = "Updated margin projection";
    const final = await PresentationFile.exportPptx(imported);
    await final.save(temporary[0]);
    const outputBytes = await fs.readFile(temporary[0]);
    const outputInventory = await trendlineInventory(outputBytes);
    assert.ok(outputInventory[0].names.includes("Updated pipeline projection"));
    assert.ok(outputInventory[1].names.includes("Updated margin projection"));

    const roundTrip = await PresentationFile.importPptx(new FileBlob(outputBytes, { type: PPTX_MIME, name: path.basename(output) }));
    assert.equal(chartByName(roundTrip, "pipeline-trend").series[0].trendlines[0].forward, 1.5);
    assert.equal(chartByName(roundTrip, "pipeline-trend").series[0].trendlines[0].line.fill, "#0EA5E9");
    assert.equal(chartByName(roundTrip, "revenue-margin-trend").series[1].trendlines[0].name, "Updated margin projection");
    const verification = roundTrip.verify({ visualQa: true });
    assert.equal(verification.ok, true, verification.ndjson);
    const inspect = roundTrip.inspect({ kind: "slide,chart", maxChars: 20_000 });
    assert.match(inspect.ndjson, /pipeline-trend/);

    const rendered = await renderArtifact(roundTrip, {
      slide: roundTrip.slides.getItem(0),
      format: "png",
      renderer: playwrightRenderer,
      viewport: { width: 1280, height: 720 },
    });
    const previewBytes = new Uint8Array(await rendered.arrayBuffer());
    if (previewBytes.byteLength < 1_000) throw new Error("Trendline preview is unexpectedly empty.");
    await fs.writeFile(temporary[1], previewBytes);

    const audit = {
      schema: "office-kit.pptx-audit.v1",
      status: "succeeded",
      provider: { actual: "office-kit", silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: { type: "greenfield-native-chart-trendline-author-edit", chartTypes: ["line", "combo"], trendlineTypes: ["linear", "movingAvg", "poly", "exp"] },
      source: { kind: "in-memory-presentation", firstExportSha256: sha256(firstBytes), bytes: firstBytes.byteLength },
      output: { path: output, sha256: sha256(outputBytes), bytes: outputBytes.byteLength },
      preview: { path: preview, sha256: sha256(previewBytes), bytes: previewBytes.byteLength, renderer: "model-svg+playwright" },
      validation: { verify: { ok: true }, inspect: { ok: true, chartCount: 2 }, package: { ok: true, charts: outputInventory }, reimport: { ok: true } },
      warnings: ["Trendline labels and complex native line graphs remain source-owned and fail closed on edit."],
    };
    await fs.writeFile(temporary[2], JSON.stringify(audit, null, 2));
    await fs.rename(temporary[0], output);
    await fs.rename(temporary[1], preview);
    await fs.rename(temporary[2], auditPathname);
    return { outputPath: output, previewPath: preview, auditPath: auditPathname, audit };
  } catch (error) {
    await Promise.all(temporary.map((entry) => fs.rm(entry, { force: true })));
    throw error;
  }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [outputPath = "output/chart-trendlines.pptx", previewPath = "output/chart-trendlines.png", auditPath = "output/chart-trendlines.audit.json"] = process.argv.slice(2);
  const result = await createAndEditTrendlineDeck({ outputPath, previewPath, auditPath });
  console.log(JSON.stringify({ outputPath: result.outputPath, previewPath: result.previewPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
