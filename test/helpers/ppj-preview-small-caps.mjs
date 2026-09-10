import assert from "node:assert/strict";
import path from "node:path";
import { writeFile } from "node:fs/promises";
import sharp from "sharp";
import JSZip from "jszip";

// Independent SVG font-backend reference: no painter output is used to build
// the expected text, coordinates, font or variant. This verifies local review
// rendering, not PowerPoint small-cap synthesis or font substitution fidelity.
export async function checkNativeSmallCaps({ pairBase, sourceWorkspace, compilePpjWorkspace,
  projectPptxToPpj, withoutAuthoredSnapshot, savePaint, assertProductionEntry, artifacts, sha256 }) {
  const cases = [], failures = [];
  const profiles = [
    { name: "run" }, { name: "default", inherited: true },
    { name: "override-none", inherited: true, caps: "none" },
    { name: "override-all", inherited: true, caps: "all" },
    { name: "shape", owner: "shape" }, { name: "table", owner: "table" },
    { name: "bold-italic", bold: true, italic: true },
    { name: "zero-alpha", alpha: 0 }, { name: "letter-spacing", spacing: 3 },
    { name: "turkish", language: "tr-TR", literal: "Ii ıİ" },
    { name: "line-break", literal: "Hh\nFf" }, { name: "lowercase", literal: "hello" },
  ];
  const esc = s => String(s).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");
  const reference = (profile, caps = profile.caps ?? "small", literal = profile.literal ?? "Hh Aa") => {
    const language = profile.language ?? "en-US";
    const displayed = caps === "all" ? literal.toLocaleUpperCase(language) : literal;
    const lines = displayed.split("\n").map((line, i) => `<text x="100" y="${140 + i * 48}" xml:space="preserve"><tspan font-family="Liberation Sans" font-size="40" font-weight="${profile.bold ? "bold" : "normal"}" font-style="${profile.italic ? "italic" : "normal"}" fill="#000000" fill-opacity="${profile.alpha ?? 1}" letter-spacing="${profile.spacing ?? 0}"${caps === "small" ? ` font-variant="small-caps" xml:lang="${language}"` : ""}>${esc(line)}</tspan></text>`).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="500" height="250"><rect width="500" height="250" fill="white"/>${lines}</svg>`;
  };
  const pixels = svg => sharp(Buffer.from(svg)).extract({ left: 90, top: 90, width: 310, height: 130 })
    .flatten({ background: "#FFFFFF" }).ensureAlpha().raw().toBuffer();
  const programFor = profile => {
    const program = structuredClone(pairBase), caps = profile.caps ?? "small";
    const style = { size: 40, fontFamily: "Liberation Sans", language: profile.language ?? "en-US",
      color: "#000000", bold: profile.bold ?? false, italic: profile.italic ?? false,
      letterSpacing: profile.spacing ?? 0, ...(profile.alpha === 0 ? { color: "#00000000" } : {}) };
    if (!profile.inherited || profile.caps) style.capitalization = caps;
    const paragraph = { style: { alignment: "left", ...(profile.inherited ? { defaultText: { capitalization: "small" } } : {}) },
      runs: [{ text: profile.literal ?? "Hh Aa", style }] };
    const text = { paragraphs: [paragraph] }, box = { x: 100, y: 100, width: 300, height: 120 };
    const margins = { left: 0, right: 0, top: 0, bottom: 0 };
    const element = profile.owner === "table"
      ? { id: "caps", type: "table", frame: box, columns: [{ width: 300 }], rows: [{ height: 120,
        cells: [{ text: { ...text, style: { margins } }, fill: { type: "solid", color: "#FFFFFF" } }] }] }
      : { id: "caps", type: profile.owner ?? "text", frame: box, text,
        ...(profile.owner === "shape"
          ? { style: { fill: { type: "solid", color: "#FFFFFF" } }, textStyle: { margins } }
          : { style: { margins } }),
        ...(profile.owner === "shape" ? { geometry: { kind: "preset", preset: "rect" } } : {}) };
    program.pages[0].background = { type: "solid", color: "#FFFFFF" };
    program.pages[0].elements = [element, { id: "control", type: "shape", frame: { x: 500, y: 250, width: 30, height: 30 },
      geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } }];
    return program;
  };
  const check = async (name, profile, workspace, receipt, caps = profile.caps ?? "small", literal = profile.literal ?? "Hh Aa") => {
    const candidateFile = `small-caps-${name}.pptx`, inputFile = `small-caps-${name}.ppj`;
    await writeFile(path.join(artifacts, candidateFile), receipt.file, { flag: "wx" });
    await writeFile(path.join(artifacts, inputFile), workspace.program, { flag: "wx" });
    const element = receipt.previewScene.presentation.slides[0].elements[0].content.value;
    const owner = profile.owner === "table" ? element.rows[0].cells[0] : element;
    if (owner.textBody) {
      const paragraph = owner.textBody.paragraphs[0], run = paragraph.runs[0];
      assert.equal(run.content.value, literal);
      assert.equal(run.fontCaps ?? paragraph.defaultRunStyle?.value?.fontCaps, caps);
    } else {
      // The importer represents uniform cells as literal text plus direct
      // textStyle, not a rich body. Assert that real native state unchanged.
      assert.equal(profile.owner, "table");
      assert.equal(owner.text, literal);
      assert.equal(owner.textStyle?.fontCaps, caps);
    }
    const ref = reference(profile, caps, literal), referenceFile = `small-caps-${name}-reference.svg`;
    await writeFile(path.join(artifacts, referenceFile), ref, { flag: "wx" });
    const painted = await savePaint(`small-caps-${name}`, receipt);
    assert.deepEqual(await pixels(painted.pages[0].svg), await pixels(ref), `${name}: independent local SVG reference`);
    assert.equal(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-small-caps-review"), caps === "small");
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-capitalization"));
    const rgba = await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: 510, top: 260, width: 1, height: 1 }).ensureAlpha().raw().toBuffer();
    assert.deepEqual([...rgba], [204, 85, 0, 255]);
    if (caps === "small" && profile.alpha !== 0) {
      assert.notDeepEqual(await pixels(ref), await pixels(reference(profile, "none", literal)), "small capitals must actually change lowercase ink");
      assert.notDeepEqual(await pixels(ref), await pixels(reference(profile, "all", literal)), "small capitals must not become full-size capitals");
    }
    const publication = await assertProductionEntry(`small-caps-${name}`, workspace, receipt, painted);
    assert.equal(publication.reliability.status, "requires-review");
    cases.push({ name, caps, literalPreserved: true, independentRasterEqual: true, foregroundControlPreserved: true,
      candidateFile, candidateSha256: sha256(receipt.file), inputFile, inputSha256: sha256(workspace.program),
      referenceFile, referenceSha256: sha256(Buffer.from(ref)) });
  };
  for (const profile of profiles) try {
    const program = programFor(profile), workspace = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) };
    const inputHash = sha256(workspace.program);
    const authored = await compilePpjWorkspace(workspace, { includePreviewScene: true });
    await check(`${profile.name}-author`, profile, workspace, authored);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const sourceFile = `small-caps-${profile.name}-original.pptx`;
    await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
    const projection = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
    const bound = { program: projection.programJson, source, assets: projection.assets };
    const noop = await compilePpjWorkspace(bound, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    await check(`${profile.name}-source`, profile, bound, noop);
    Object.assign(cases.at(-1), { sourceFile, sourceSha256: sourceHash, exactNoop: true });
    if (profile.name === "run") for (const edit of ["none", "all", "text"]) try {
      const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
      const leaf = request.pages[0].elements[0].nativeRef.leaves.find(l => l.kind === (edit === "text" ? "text" : "fontCaps"));
      assert.ok(leaf, `${edit}: issued leaf`);
      leaf.value = edit === "text" ? "Ff Ii" : edit;
      const changed = { ...bound, program: Buffer.from(JSON.stringify(request)) }, requestHash = sha256(changed.program);
      const requestFile = `small-caps-edit-${edit}.request.ppj`;
      await writeFile(path.join(artifacts, requestFile), changed.program, { flag: "wx" });
      const candidate = await compilePpjWorkspace(changed, { includePreviewScene: true });
      await check(`edit-${edit}`, profile, changed, candidate, edit === "text" ? "small" : edit, edit === "text" ? "Ff Ii" : "Hh Aa");
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `small-caps-edited-${edit}.pptx`, assetRootUri: "assets" });
      const actual = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements[0].text.paragraphs[0].runs[0];
      assert.equal(actual.text, edit === "text" ? "Ff Ii" : "Hh Aa");
      assert.equal(actual.style.capitalization, edit === "text" ? "small" : edit);
      const old = await JSZip.loadAsync(source), updated = await JSZip.loadAsync(candidate.file), parts = [];
      assert.deepEqual(Object.keys(updated.files).sort(), Object.keys(old.files).sort());
      for (const name of Object.keys(old.files)) if (!old.files[name].dir &&
        !Buffer.from(await old.file(name).async("uint8array")).equals(Buffer.from(await updated.file(name).async("uint8array")))) parts.push(name);
      assert.deepEqual(parts, ["ppt/slides/slide1.xml"]);
      const beforeXml = await old.file(parts[0]).async("string"), afterXml = await updated.file(parts[0]).async("string");
      assert.equal(afterXml, edit === "text" ? beforeXml.replace(">Hh Aa<", ">Ff Ii<") : beforeXml.replace('cap="small"', `cap="${edit}"`));
      assert.deepEqual((await compilePpjWorkspace(changed)).file, candidate.file, "read-only scene option leaves candidate identical");
      assert.equal(sha256(changed.program), requestHash); assert.equal(sha256(source), sourceHash);
      const reprojectionFile = `small-caps-edit-${edit}.reprojected.ppj`;
      await writeFile(path.join(artifacts, reprojectionFile), fresh.programJson, { flag: "wx" });
      Object.assign(cases.at(-1), { sourceFile, sourceSha256: sourceHash, requestFile, requestSha256: requestHash,
        reprojectionFile, reprojectionSha256: sha256(fresh.programJson), changedParts: parts, nonTargetXmlPreserved: true });
    } catch (error) { failures.push({ name: `edit-${edit}`, code: error.code, message: error.message }); }
    assert.equal(sha256(source), sourceHash); assert.equal(sha256(workspace.program), inputHash);
  } catch (error) { failures.push({ name: profile.name, code: error.code, message: error.message }); }
  return { cases, failures, expectedCount: profiles.length * 2 + 3 };
}
