import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "office-kit-pack-"));

try {
  const packed = run("npm", ["pack", repoRoot, "--json", "--ignore-scripts", "--pack-destination", temporary], repoRoot);
  const report = JSON.parse(packed.stdout)[0];
  const tarball = path.join(temporary, report.filename);
  assert.ok(fs.existsSync(tarball), `npm pack did not create ${tarball}`);
  const dependencyTarballs = packProductionDependencies(temporary);
  // Exercise a real npm install without making a release gate depend on the
  // registry. Optional renderer peers are intentionally outside this core
  // OfficeKit/PDF probe and remain covered by package metadata tests.
  run("npm", [
    "install", "--offline", "--ignore-scripts", "--no-audit", "--no-fund",
    "--omit=dev", "--legacy-peer-deps", "--no-save", tarball, ...dependencyTarballs,
  ], temporary);
  testGlobalCli({ dependencyTarballs, tarball, temporary });

  const probe = String.raw`
    import { spawnSync } from "node:child_process";
    import { createHash } from "node:crypto";
    import fs from "node:fs";
    import path from "node:path";
    import { pathToFileURL } from "node:url";

    import {
      DocumentFile, DocumentModel, FileBlob, PdfArtifact, PdfFile,
      Presentation, PresentationFile, SpreadsheetFile, Workbook,
    } from "office-kit";

    const canonicalCodec = await import("office-kit/codec");
    const canonicalWire = await import("office-kit/codec/wire");
    if (
      typeof canonicalCodec.exportXlsxWithOfficeKit !== "function" ||
      typeof canonicalCodec.importXlsxWithOfficeKit !== "function" ||
      typeof canonicalCodec.invokeOfficeKit !== "function" ||
      typeof canonicalCodec.officeKitStatus !== "function" ||
      canonicalWire.CodecRequestSchema == null ||
      canonicalWire.CodecResponseSchema == null
    ) process.exit(59);

    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:B2").values = [["Label", "Value"], ["clean install", 7]];
    sheet.getRange("D1:E3").values = [["X", "Y"], [1, 3], [2, 8]];
    const scatter = sheet.charts.add("scatter", sheet.getRange("D1:E3"));
    scatter.title = "Packed scatter";
    scatter.series.items[0].marker = { symbol: "circle", size: 6, fill: "#0EA5E9" };
    sheet.getRange("G1:I3").values = [["X", "Y", "Size"], [1, 3, 4], [2, 8, 9]];
    const bubble = sheet.charts.add("bubble", sheet.getRange("G1:I3"));
    bubble.title = "Packed bubble";
    bubble.series.items[0].fill = "#38BDF8";
    const xlsx = await SpreadsheetFile.exportXlsx(workbook);
    if (xlsx.metadata.codec !== "office-kit" || xlsx.bytes[0] !== 0x50 || xlsx.bytes[1] !== 0x4b) process.exit(1);
    const importedWorkbook = await SpreadsheetFile.importXlsx(xlsx);
    if (importedWorkbook.worksheets.getItem("Packaged").getRange("B2").values[0][0] !== 7) process.exit(2);
    const importedScatter = importedWorkbook.worksheets.getItem("Packaged").charts.items[0];
    if (importedScatter.type !== "scatter" || importedScatter.xAxis.axisType !== "valueAxis") process.exit(4);
    if (JSON.stringify(importedScatter.series.items[0].xValues) !== "[1,2]") process.exit(5);
    const importedBubble = importedWorkbook.worksheets.getItem("Packaged").charts.items[1];
    if (importedBubble.type !== "bubble" || importedBubble.xAxis.axisType !== "valueAxis") process.exit(6);
    if (JSON.stringify(importedBubble.series.items[0].bubbleSizes) !== "[4,9]") process.exit(7);
    const xlsx2 = await SpreadsheetFile.exportXlsx(importedWorkbook, { recalculate: false });
    if ((await SpreadsheetFile.importXlsx(xlsx2)).worksheets.getItem("Packaged").getRange("A2").values[0][0] !== "clean install") process.exit(3);
    {
      const packagedRoot = path.join(process.cwd(), "node_modules", "office-kit");
      const validationWorkflowPath = path.join(
        packagedRoot,
        "skills", "spreadsheets", "skills", "spreadsheets", "examples", "officekit-data-validation-workflow.mjs",
      );
      if (!fs.existsSync(validationWorkflowPath)) process.exit(8);
      const validationOutput = path.join(process.cwd(), "packed-data-validation.xlsx");
      const { createDataValidationWorkbook } = await import(pathToFileURL(validationWorkflowPath).href);
      const validationResult = await createDataValidationWorkbook(validationOutput);
      const validationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(validationOutput));
      const validationSheet = validationRoundTrip.worksheets.getItem("Intake");
      if (
        validationResult.audit.provider.actual !== "office-kit" ||
        validationResult.audit.provider.fallbackUsed ||
        validationSheet.dataValidations.items.length !== 3 ||
        validationSheet.dataValidations.items[0].rule.prompt !== "Pick the current workflow state." ||
        validationSheet.dataValidations.items[0].rule.errorStyle !== "information" ||
        validationSheet.dataValidations.items[0].rule.showDropdown !== false
      ) process.exit(9);
    }

    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    document.addInsertion("packaged accepted insertion", { author: "Package QA" });
    document.addDeletion("packaged removed deletion", { author: "Package QA" });
    document.setSettings({ trackRevisions: true, documentProtection: "comments" });
    document.addWatermark("PACKAGED DRAFT", { sectionIndex: 0 });
    document.addImage({
      dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==",
      alt: "Packed floating image",
      widthPx: 80,
      heightPx: 60,
      placement: {
        type: "floating",
        horizontal: { relativeTo: "margin", offsetPx: 24 },
        vertical: { relativeTo: "paragraph", offsetPx: 0 },
        wrap: "square",
        wrapSide: "right",
        distanceFromTextPx: { top: 2, right: 8, bottom: 2, left: 8 },
      },
    });
    const docx = await DocumentFile.exportDocx(document);
    if (docx.metadata.codec !== "office-kit" || docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(10);
    const importedDocument = await DocumentFile.importDocx(docx);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(11);
    if (importedDocument.settings.documentProtection?.edit !== "comments") process.exit(42);
    if (importedDocument.watermarks.length !== 1 || importedDocument.watermarks[0].text !== "PACKAGED DRAFT") process.exit(46);
    const importedImage = importedDocument.blocks.find((block) => block.kind === "image");
    if (
      importedImage?.placement?.type !== "floating" ||
      importedImage.placement.horizontal.relativeTo !== "margin" ||
      importedImage.placement.vertical.relativeTo !== "paragraph" ||
      importedImage.placement.wrap !== "square" ||
      importedImage.placement.wrapSide !== "right"
    ) process.exit(44);
    importedImage.placement = {
      type: "floating",
      horizontal: { relativeTo: "page", offsetPx: 36 },
      vertical: { relativeTo: "paragraph", offsetPx: 0 },
      wrap: "topAndBottom",
      distanceFromTextPx: { top: 4, right: 0, bottom: 4, left: 0 },
    };
    importedDocument.watermarks[0].text = "PACKAGED REVIEW";
    const packagedDocument2 = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDocument));
    if (packagedDocument2.blocks[0].text !== "clean install DOCX") process.exit(12);
    const packagedImage2 = packagedDocument2.blocks.find((block) => block.kind === "image");
    if (packagedImage2?.placement?.wrap !== "topAndBottom" || packagedImage2.placement.horizontal.relativeTo !== "page") process.exit(45);
    if (packagedDocument2.watermarks.length !== 1 || packagedDocument2.watermarks[0].text !== "PACKAGED REVIEW") process.exit(47);
    const docxSourceHash = createHash("sha256").update(docx.bytes).digest("hex");
    const finalizedDocx = await DocumentFile.finalizeRevisions(docx, {
      mode: "accept",
      expectedSourceSha256: docxSourceHash,
    });
    const finalization = finalizedDocx.metadata.revisionFinalization;
    if (finalization.sourceSha256 !== docxSourceHash || finalization.insertionCount !== 1 || finalization.deletionCount !== 1) process.exit(13);
    if (finalization.trackingBefore !== true || finalization.trackingAfter !== false) process.exit(14);
    if (JSON.stringify(finalization.changedParts) !== JSON.stringify(["word/document.xml", "word/settings.xml"])) process.exit(15);
    const finalizedDocument = await DocumentFile.importDocx(finalizedDocx);
    if (finalizedDocument.blocks.some((block) => block.kind === "change") || finalizedDocument.settings.trackRevisions) process.exit(16);
    if (finalizedDocument.settings.documentProtection?.edit !== "comments") process.exit(43);
    if (!finalizedDocument.blocks.some((block) => block.text === "packaged accepted insertion") || finalizedDocument.blocks.some((block) => block.text === "packaged removed deletion")) process.exit(17);
    if (createHash("sha256").update(docx.bytes).digest("hex") !== docxSourceHash) process.exit(18);

    const packageRoot = path.join(process.cwd(), "node_modules", "office-kit");
    const lineNumberingWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-section-line-numbering-edit-workflow.mjs",
    );
    if (!fs.existsSync(lineNumberingWorkflowPath)) process.exit(63);
    const lineNumberingSource = DocumentModel.create({ blocks: [] });
    lineNumberingSource.addParagraph("Packaged source-bound line-numbering transaction.");
    lineNumberingSource.addSection({
      breakType: "nextPage",
      lineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
    });
    lineNumberingSource.addParagraph("Only this section's line-number settings may change.");
    const lineNumberingInput = path.join(process.cwd(), "packed-line-numbering-input.docx");
    const lineNumberingOutput = path.join(process.cwd(), "packed-line-numbering-output.docx");
    const lineNumberingAudit = path.join(process.cwd(), "packed-line-numbering-audit.json");
    await (await DocumentFile.exportDocx(lineNumberingSource)).save(lineNumberingInput);
    const { editImportedSectionLineNumbering } = await import(pathToFileURL(lineNumberingWorkflowPath).href);
    const lineNumberingResult = await editImportedSectionLineNumbering({
      inputPath: lineNumberingInput,
      outputPath: lineNumberingOutput,
      auditPath: lineNumberingAudit,
      sectionBlockIndex: 1,
      expectedLineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
      replacementLineNumbering: { countBy: 10, start: 4, distance: 480, restart: "continuous" },
    });
    const lineNumberingRoundTrip = await DocumentFile.importDocx(await FileBlob.load(lineNumberingOutput));
    if (
      lineNumberingResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(lineNumberingResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      JSON.stringify(lineNumberingRoundTrip.blocks[1]?.lineNumbering) !== JSON.stringify({ countBy: 10, start: 4, distance: 480, restart: "continuous" })
    ) process.exit(64);

    const columnsWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-section-columns-edit-workflow.mjs",
    );
    if (!fs.existsSync(columnsWorkflowPath)) process.exit(65);
    const columnsSource = DocumentModel.create({ blocks: [] });
    columnsSource.addParagraph("Packaged source-bound section-columns transaction.");
    columnsSource.addSection({
      breakType: "nextPage",
      columns: { count: 2, spacing: 720, separator: true },
    });
    columnsSource.addParagraph("Only this section's column profile may change.");
    const columnsInput = path.join(process.cwd(), "packed-columns-input.docx");
    const columnsOutput = path.join(process.cwd(), "packed-columns-output.docx");
    const columnsAudit = path.join(process.cwd(), "packed-columns-audit.json");
    await (await DocumentFile.exportDocx(columnsSource)).save(columnsInput);
    const { editImportedSectionColumns } = await import(pathToFileURL(columnsWorkflowPath).href);
    const columnsResult = await editImportedSectionColumns({
      inputPath: columnsInput,
      outputPath: columnsOutput,
      auditPath: columnsAudit,
      sectionBlockIndex: 1,
      expectedColumns: { count: 2, spacing: 720, separator: true },
      replacementColumns: { count: 3, spacing: 360, separator: false },
    });
    const columnsRoundTrip = await DocumentFile.importDocx(await FileBlob.load(columnsOutput));
    if (
      columnsResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(columnsResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      JSON.stringify(columnsRoundTrip.blocks[1]?.columns) !== JSON.stringify({ count: 3, spacing: 360, separator: false })
    ) process.exit(66);

    const sectionBreakWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-section-break-edit-workflow.mjs",
    );
    if (!fs.existsSync(sectionBreakWorkflowPath)) process.exit(67);
    const sectionBreakSource = DocumentModel.create({ blocks: [] });
    sectionBreakSource.addParagraph("Packaged source-bound section-break transaction.");
    sectionBreakSource.addSection({ breakType: "nextPage" });
    sectionBreakSource.addParagraph("Only this section's break type may change.");
    const sectionBreakInput = path.join(process.cwd(), "packed-section-break-input.docx");
    const sectionBreakOutput = path.join(process.cwd(), "packed-section-break-output.docx");
    const sectionBreakAudit = path.join(process.cwd(), "packed-section-break-audit.json");
    await (await DocumentFile.exportDocx(sectionBreakSource)).save(sectionBreakInput);
    const { editImportedSectionBreakType } = await import(pathToFileURL(sectionBreakWorkflowPath).href);
    const sectionBreakResult = await editImportedSectionBreakType({
      inputPath: sectionBreakInput,
      outputPath: sectionBreakOutput,
      auditPath: sectionBreakAudit,
      sectionBlockIndex: 1,
      expectedBreakType: "nextPage",
      replacementBreakType: "continuous",
    });
    const sectionBreakRoundTrip = await DocumentFile.importDocx(await FileBlob.load(sectionBreakOutput));
    if (
      sectionBreakResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(sectionBreakResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      sectionBreakRoundTrip.blocks[1]?.breakType !== "continuous"
    ) process.exit(68);

    const tableColumnWidthsWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-table-column-widths-edit-workflow.mjs",
    );
    if (!fs.existsSync(tableColumnWidthsWorkflowPath)) process.exit(69);
    const tableColumnWidthsSource = DocumentModel.create({ blocks: [] });
    tableColumnWidthsSource.addParagraph("Packaged source-bound table column-width transaction.");
    tableColumnWidthsSource.addTable({
      values: [["Quarter", "Revenue", "Margin"], ["Q1", "1.2M", "44%"]],
      widthDxa: 9300,
      indentDxa: 120,
      columnWidthsDxa: [2100, 4500, 2700],
      cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
      borderColor: "445566",
      borderSize: 8,
      headerFill: "E2E8F0",
    });
    const tableColumnWidthsInput = path.join(process.cwd(), "packed-table-column-widths-input.docx");
    const tableColumnWidthsOutput = path.join(process.cwd(), "packed-table-column-widths-output.docx");
    const tableColumnWidthsAudit = path.join(process.cwd(), "packed-table-column-widths-audit.json");
    await (await DocumentFile.exportDocx(tableColumnWidthsSource)).save(tableColumnWidthsInput);
    const { editImportedTableColumnWidths } = await import(pathToFileURL(tableColumnWidthsWorkflowPath).href);
    const tableColumnWidthsResult = await editImportedTableColumnWidths({
      inputPath: tableColumnWidthsInput,
      outputPath: tableColumnWidthsOutput,
      auditPath: tableColumnWidthsAudit,
      tableBlockIndex: 1,
      expectedColumnWidthsDxa: [2100, 4500, 2700],
      replacementColumnWidthsDxa: [3000, 3600, 2700],
    });
    const tableColumnWidthsRoundTrip = await DocumentFile.importDocx(await FileBlob.load(tableColumnWidthsOutput));
    if (
      tableColumnWidthsResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(tableColumnWidthsResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      JSON.stringify(tableColumnWidthsRoundTrip.blocks[1]?.columnWidthsDxa) !== JSON.stringify([3000, 3600, 2700])
    ) process.exit(70);

    const tableFormattingWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-table-formatting-edit-workflow.mjs",
    );
    if (!fs.existsSync(tableFormattingWorkflowPath)) process.exit(71);
    const tableFormattingOutput = path.join(process.cwd(), "packed-table-formatting-output.docx");
    const tableFormattingAudit = path.join(process.cwd(), "packed-table-formatting-audit.json");
    const { editImportedTableFormatting } = await import(pathToFileURL(tableFormattingWorkflowPath).href);
    const tableFormattingResult = await editImportedTableFormatting({
      inputPath: tableColumnWidthsInput,
      outputPath: tableFormattingOutput,
      auditPath: tableFormattingAudit,
      tableBlockIndex: 1,
      expectedFormatting: {
        indentDxa: 120,
        cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
        borderColor: "445566",
        borderSize: 8,
        headerFill: "E2E8F0",
      },
      replacementFormatting: {
        indentDxa: 240,
        cellMarginsDxa: { top: 100, bottom: 120, start: 160, end: 180 },
        borderColor: "224466",
        borderSize: 12,
        headerFill: "DDEBF7",
      },
    });
    const tableFormattingRoundTrip = await DocumentFile.importDocx(await FileBlob.load(tableFormattingOutput));
    if (
      tableFormattingResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(tableFormattingResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      tableFormattingRoundTrip.blocks[1]?.indentDxa !== 240 ||
      tableFormattingRoundTrip.blocks[1]?.headerFill !== "DDEBF7"
    ) process.exit(72);

    const imageAltTextWorkflowPath = path.join(
      packageRoot,
      "skills", "documents", "skills", "documents", "examples", "officekit-image-alt-text-edit-workflow.mjs",
    );
    if (!fs.existsSync(imageAltTextWorkflowPath)) process.exit(73);
    const imageAltTextSource = DocumentModel.create({ blocks: [] });
    imageAltTextSource.addParagraph("Packaged source-bound image alternative-text transaction.");
    imageAltTextSource.addImage({
      dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==",
      alt: "Packaged source chart",
      widthPx: 48,
      heightPx: 36,
      placement: {
        type: "floating",
        horizontal: { relativeTo: "margin", offsetPx: 24 },
        vertical: { relativeTo: "paragraph", offsetPx: 0 },
        wrap: "square",
        wrapSide: "bothSides",
        distanceFromTextPx: { top: 0, right: 4, bottom: 0, left: 4 },
      },
    });
    const imageAltTextInput = path.join(process.cwd(), "packed-image-alt-text-input.docx");
    const imageAltTextOutput = path.join(process.cwd(), "packed-image-alt-text-output.docx");
    const imageAltTextAudit = path.join(process.cwd(), "packed-image-alt-text-audit.json");
    await (await DocumentFile.exportDocx(imageAltTextSource)).save(imageAltTextInput);
    const { editImportedImageAltText } = await import(pathToFileURL(imageAltTextWorkflowPath).href);
    const imageAltTextResult = await editImportedImageAltText({
      inputPath: imageAltTextInput,
      outputPath: imageAltTextOutput,
      auditPath: imageAltTextAudit,
      imageBlockIndex: 1,
      expectedAlt: "Packaged source chart",
      replacementAlt: "Packaged accessible chart description",
    });
    const imageAltTextRoundTrip = await DocumentFile.importDocx(await FileBlob.load(imageAltTextOutput));
    if (
      imageAltTextResult.audit.provider.actual !== "office-kit" ||
      JSON.stringify(imageAltTextResult.audit.validation.changedParts) !== JSON.stringify(["word/document.xml"]) ||
      imageAltTextRoundTrip.blocks[1]?.alt !== "Packaged accessible chart description"
    ) process.exit(74);

    const presentation = Presentation.create();
    presentation.slides.add({ name: "Packaged" }).shapes.add({
      name: "Title", geometry: "roundRect", text: "clean install PPTX",
      position: { left: 40, top: 40, width: 520, height: 80 },
    });
    const pptx = await PresentationFile.exportPptx(presentation);
    if (pptx.metadata.codec !== "office-kit" || pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(20);
    const importedPresentation = await PresentationFile.importPptx(pptx);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(21);
    if ((await PresentationFile.importPptx(await PresentationFile.exportPptx(importedPresentation))).slides.count !== 1) process.exit(22);

    const installedPackage = path.join(process.cwd(), "node_modules", "office-kit");
    const installedBin = path.join(
      process.cwd(),
      "node_modules",
      ".bin",
      process.platform === "win32" ? "officekit.cmd" : "officekit",
    );
    if (!fs.existsSync(installedBin)) process.exit(62);
    const initializedProject = path.join(process.cwd(), "officekit-initialized-project");
    const initialized = spawnSync(
      installedBin,
      [
        "init",
        initializedProject,
        "--tools",
        "agents",
        "--json",
      ],
      {
        cwd: process.cwd(),
        encoding: "utf8",
        shell: process.platform === "win32",
      },
    );
    if (initialized.status !== 0) {
      process.stderr.write(initialized.stderr);
      process.exit(60);
    }
    const initializedResult = JSON.parse(initialized.stdout);
    if (
      initializedResult.created !== 7 ||
      initializedResult.tools[0]?.id !== "agents" ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "office-kit", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "documents", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "spreadsheets", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "excel-live-control", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "presentations", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "pdf", "SKILL.md")) ||
      !fs.existsSync(path.join(initializedProject, ".agents", "skills", "template-creator", "SKILL.md"))
    ) process.exit(61);

    const duplicateWorkflowPath = path.join(
      installedPackage,
      "skills", "presentations", "skills", "presentations", "examples", "officekit-slide-duplicate-workflow.mjs",
    );
    if (!fs.existsSync(duplicateWorkflowPath)) process.exit(23);
    const cloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
    const cloneSource = cloneFixture.slides.add({
      name: "Packed clone source",
      notes: "Packaged closed-leaf clone notes.",
    });
    const cloneGroup = cloneSource.addGroup({
      name: "packed-cluster",
      position: { left: 48, top: 40, width: 320, height: 120 },
      childFrame: { left: 0, top: 0, width: 320, height: 120 },
    });
    const cloneLeft = cloneGroup.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
    const cloneRight = cloneGroup.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
    cloneGroup.connectors.add({
      name: "join", from: cloneLeft, to: cloneRight,
      start: { x: 90, y: 41 }, end: { x: 210, y: 41 }, line: { fill: "#64748B", width: 1 },
    });
    cloneSource.shapes.add({
      name: "packed-clone-links",
      geometry: "textbox",
      position: { left: 48, top: 190, width: 420, height: 64 },
      fill: "transparent",
      line: { fill: "transparent", width: 0 },
      text: [{ runs: [
        { text: "Guide ", link: { uri: "https://example.com/packed-clone" } },
        { text: "Next ", link: { action: "nextSlide" } },
        { text: "Review route", link: { customShow: "Packed route", returnToSlide: true } },
      ] }],
    });
    cloneFixture.customShows.add({ name: "Packed route", nativeId: 23, slides: [cloneSource] });
    cloneSource.comments.addThread(undefined, "Packaged closed-leaf clone comment.", {
      author: "Package QA",
      created: "2026-07-18T03:05:00Z",
      position: { x: 360, y: 240 },
    });
    const cloneInput = path.join(process.cwd(), "packed-clone-source.pptx");
    const cloneOutput = path.join(process.cwd(), "packed-clone-output.pptx");
    const cloneAudit = path.join(process.cwd(), "packed-clone-audit.json");
    await (await PresentationFile.exportPptx(cloneFixture)).save(cloneInput);
    const { duplicatePptxSlide } = await import(pathToFileURL(duplicateWorkflowPath).href);
    const cloneResult = await duplicatePptxSlide({
      inputPath: cloneInput,
      outputPath: cloneOutput,
      auditPath: cloneAudit,
      expectedName: "Packed clone source",
      allowClosedLeaves: true,
    });
    if (
      cloneResult.audit.operation.clonePart !== "ppt/slides/slide2.xml" ||
      cloneResult.audit.operation.runHyperlinks.relationshipCount !== 1 ||
      cloneResult.audit.operation.runHyperlinks.actionOnlyCount !== 2 ||
      cloneResult.audit.operation.runHyperlinks.customShowCount !== 1 ||
      !cloneResult.audit.validation.package.runHyperlinks.exactSourceGraphRetained ||
      !cloneResult.audit.validation.package.customShows.exactSourceMembershipRetained ||
      !cloneResult.audit.validation.reimport.customShowMembershipRetained ||
      !cloneResult.audit.operation.closedLeaves.speakerNotes ||
      !cloneResult.audit.operation.closedLeaves.legacyComments ||
      !cloneResult.audit.validation.package.retainedSourcePartsByteIdentical ||
      !cloneResult.audit.validation.package.closedLeaves.speakerNotes?.notesXmlByteIdentical ||
      !cloneResult.audit.validation.package.closedLeaves.legacyComments?.commentsXmlByteIdentical ||
      !cloneResult.audit.validation.reimport.sourceAndCloneSemanticsEqual ||
      !cloneResult.audit.validation.reimport.sourceAndCloneClosedLeavesEqual ||
      !cloneResult.audit.validation.modelRender.visualEquivalent
    ) process.exit(24);
    const packedClone = await PresentationFile.importPptx(new FileBlob(await fs.promises.readFile(cloneOutput), {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      name: "packed-clone-output.pptx",
    }));
    if (
      packedClone.slides.count !== 2 ||
      packedClone.slides.getItem(1).groups.items[0].connectors.items[0].startTargetId !== packedClone.slides.getItem(1).groups.items[0].shapes.items[0].id ||
      packedClone.slides.getItem(1).speakerNotes.text !== "Packaged closed-leaf clone notes." ||
      packedClone.slides.getItem(1).comments.items[0].comments[0].text !== "Packaged closed-leaf clone comment." ||
      packedClone.slides.getItem(1).shapes.items.find((shape) => shape.name === "packed-clone-links").text.paragraphs[0].runs[2].link.customShow !== "Packed route" ||
      JSON.stringify(packedClone.customShows.getItem("Packed route").slideIds) !== JSON.stringify([packedClone.slides.getItem(0).id])
    ) process.exit(25);

    const pdf = PdfArtifact.create({ pages: [{ text: "clean install PDF" }] });
    const pdfFile = await PdfFile.exportPdf(pdf);
    if (pdfFile.bytes[0] !== 0x25 || pdfFile.bytes[1] !== 0x50 || pdfFile.bytes[2] !== 0x44 || pdfFile.bytes[3] !== 0x46) process.exit(30);
    const inspection = await PdfFile.inspectPdf(pdfFile);
    if (inspection.summary.pages !== 1 || !inspection.summary.tagged) process.exit(31);
    const importedPdf = await PdfFile.importPdf(pdfFile);
    if (!importedPdf.extractText().includes("clean install PDF")) process.exit(32);
    const pyhankoProviderPath = path.join(
      installedPackage,
      "skills", "pdf", "skills", "pdf", "scripts", "pyhanko_provider.py",
    );
    if (
      !fs.existsSync(pyhankoProviderPath) ||
      !fs.readFileSync(pyhankoProviderPath, "utf8").includes("office-kit.pyhanko-verify.v1")
    ) process.exit(33);
    const verapdfProviderPath = path.join(
      installedPackage,
      "skills", "pdf", "skills", "pdf", "scripts", "verapdf_provider.py",
    );
    if (
      !fs.existsSync(verapdfProviderPath) ||
      !fs.readFileSync(verapdfProviderPath, "utf8").includes("office-kit.verapdf-validation.v1")
    ) process.exit(34);

    const creatorPath = path.join(
      installedPackage,
      "skills", "template-creator", "skills", "template-creator", "scripts", "create-template-skill.mjs",
    );
    if (!fs.existsSync(creatorPath)) process.exit(50);

    const fixtureDirectory = path.join(process.cwd(), "template-creator-fixture");
    const templateHome = path.join(process.cwd(), "template-creator-home");
    const referencePath = path.join(fixtureDirectory, "reference.xlsx");
    const previewPath = path.join(fixtureDirectory, "preview.png");
    fs.mkdirSync(fixtureDirectory, { recursive: true });
    fs.writeFileSync(referencePath, xlsx.bytes);
    fs.writeFileSync(
      previewPath,
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==",
        "base64",
      ),
    );
    const created = spawnSync(
      process.execPath,
      [
        creatorPath,
        "--reference-path", referencePath,
        "--preview-path", previewPath,
        "--display-name", "Packed workbook template",
        "--description", "Create a workbook from the clean-installed package fixture.",
      ],
      {
        cwd: process.cwd(),
        encoding: "utf8",
        env: { ...process.env, OFFICE_KIT_HOME: templateHome },
      },
    );
    if (created.status !== 0) {
      process.stderr.write(created.stderr);
      process.exit(51);
    }
    const template = JSON.parse(created.stdout);
    if (
      template.kind !== "spreadsheet" ||
      template.skillName !== "artifact-template-packed-workbook-template" ||
      path.dirname(template.skillPath) !== path.join(templateHome, "skills")
    ) process.exit(52);
    const sidecar = JSON.parse(fs.readFileSync(path.join(template.skillPath, "artifact-template.json"), "utf8"));
    const retainedReference = fs.readFileSync(path.join(template.skillPath, sidecar.reference));
    const retainedPreview = fs.readFileSync(path.join(template.skillPath, sidecar.preview));
    if (
      template.schemaVersion !== 2 ||
      sidecar.schemaVersion !== 2 ||
      sidecar.id !== "artifact-template-packed-workbook-template" ||
      sidecar.displayName !== "Packed workbook template" ||
      sidecar.kind !== "spreadsheet" ||
      sidecar.reference !== "assets/reference.xlsx" ||
      sidecar.preview !== "assets/preview.png" ||
      sidecar.visualCommitment !== "opinionated" ||
      sidecar.editProfile?.level !== "copy-only" ||
      sidecar.provenance?.referenceSha256 !== createHash("sha256").update(retainedReference).digest("hex") ||
      sidecar.provenance?.previewSha256 !== createHash("sha256").update(retainedPreview).digest("hex")
    ) process.exit(53);
    if (
      !retainedReference.equals(Buffer.from(xlsx.bytes)) ||
      !retainedPreview.equals(fs.readFileSync(previewPath))
    ) process.exit(54);

    const packagedTemplateRoot = path.join(
      installedPackage,
      "skills", "default-template-library", "skills",
    );
    if (
      !fs.existsSync(packagedTemplateRoot) ||
      fs.readdirSync(packagedTemplateRoot).filter((name) =>
        name.startsWith("artifact-template-")
      ).length !== 20
    ) process.exit(55);
    if (
      fs.existsSync(path.join(
        installedPackage,
        "skills", "office-kit", "skills", "office-kit", "scripts", "query-templates.mjs",
      ))
    ) process.exit(56);
    const queried = spawnSync(
      installedBin,
      [
        "template", "search",
        "--kind", "spreadsheet",
        "--root", path.join(templateHome, "skills"),
        "--id", template.skillName,
        "--purpose", "clean installed package fixture",
        "--json",
      ],
      {
        cwd: process.cwd(),
        encoding: "utf8",
      },
    );
    if (queried.status !== 0) {
      process.stderr.write(queried.stderr);
      process.exit(57);
    }
    const catalogResult = JSON.parse(queried.stdout);
    if (
      catalogResult.schemaVersion !== 2 ||
      catalogResult.selectionMade !== false ||
      catalogResult.ranking?.algorithm !== "bm25f" ||
      catalogResult.invalid.length !== 0 ||
      catalogResult.candidates.length !== 1 ||
      catalogResult.candidates[0].id !== template.skillName ||
      !(catalogResult.candidates[0].match?.bm25 > 0) ||
      catalogResult.candidates[0].editProfile?.level !== "copy-only" ||
      catalogResult.candidates[0].skillPath !==
        path.join(template.skillPath, "SKILL.md") ||
      catalogResult.candidates[0].referencePath !==
        path.join(template.skillPath, "assets", "reference.xlsx")
    ) process.exit(58);
  `;

  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32"
      ? `${path.dirname(process.execPath)};C:\\Windows\\System32`
      : `${path.dirname(process.execPath)}:/usr/bin:/bin`,
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("Office file Skills, PDF, OfficeKit, Template Creator, and bundled templates clean-install package smoke ok");

function run(command, args, cwd, environment = {}) {
  // npm is a .cmd shim on Windows. Invoke that shim through cmd.exe there;
  // Unix keeps the normal argv-only process invocation.
  const executable = process.platform === "win32" && command === "npm"
    ? "npm.cmd"
    : command;
  const result = spawnSync(executable, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...environment },
    // Windows .cmd shims need cmd.exe. All values here are test-owned paths
    // and fixed npm arguments; ordinary process invocations remain argv-only.
    shell: process.platform === "win32" && executable === "npm.cmd",
  });
  assert.equal(result.status, 0, `${executable} ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return result;
}

function packProductionDependencies(temporary) {
  const lock = JSON.parse(fs.readFileSync(path.join(repoRoot, "package-lock.json"), "utf8"));
  const destination = path.join(temporary, "dependency-tarballs");
  fs.mkdirSync(destination, { recursive: true });
  return Object.entries(lock.packages || {})
    .filter(([location, metadata]) => location.startsWith("node_modules/") && !metadata.dev && !metadata.optional && !metadata.peer)
    .map(([location]) => {
      const source = path.join(repoRoot, location);
      assert.ok(fs.existsSync(source), `npm ci production dependency is missing: ${location}`);
      const packed = run("npm", ["pack", source, "--json", "--ignore-scripts", "--pack-destination", destination], repoRoot);
      const report = JSON.parse(packed.stdout)[0];
      return path.join(destination, report.filename);
    });
}

function testGlobalCli({ dependencyTarballs, tarball, temporary }) {
  const globalPrefix = path.join(temporary, "global-prefix");
  run("npm", [
    "install", "--global", "--prefix", globalPrefix, "--offline",
    "--ignore-scripts", "--no-audit", "--no-fund", "--omit=dev",
    "--legacy-peer-deps", tarball, ...dependencyTarballs,
  ], temporary);
  const officekit = process.platform === "win32"
    ? path.join(globalPrefix, "officekit.cmd")
    : path.join(globalPrefix, "bin", "officekit");
  assert.ok(fs.existsSync(officekit), "global-prefix install must expose officekit");
  const globalModules = process.platform === "win32"
    ? path.join(globalPrefix, "node_modules")
    : path.join(globalPrefix, "lib", "node_modules");
  const officekitModule = path.join(globalModules, "office-kit", "bin", "officekit.mjs");
  assert.ok(fs.existsSync(officekitModule), "global-prefix install must retain the OfficeKit CLI module");
  const launcherVersion = spawnSync(officekit, ["--version"], {
    cwd: temporary,
    encoding: "utf8",
    shell: process.platform === "win32",
  });
  assert.equal(launcherVersion.status, 0, `global officekit launcher failed\nSTDERR:\n${launcherVersion.stderr}`);
  assert.equal(launcherVersion.stdout.trim(), "0.5.0");
  // A Windows .cmd launcher is intended for an interactive command shell. Node's
  // `shell: true` flattens argv before it reaches that launcher, which changes
  // values containing spaces. Check that the launcher exists, then exercise the
  // CLI module with Node's normal argv contract; the standalone Windows lane
  // separately invokes officekit.cmd from PowerShell as an end user would.
  const execute = (args, cwd, environment = {}) => spawnSync(process.execPath, [officekitModule, ...args], {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...environment },
    shell: false,
  });
  const expectSuccess = (args, cwd, environment = {}) => {
    const result = execute(args, cwd, environment);
    assert.equal(
      result.status,
      0,
      `global officekit ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
    );
    return result;
  };

  const project = path.join(temporary, "global-empty-project");
  fs.mkdirSync(project, { recursive: true });
  assert.equal(expectSuccess(["--version"], project).stdout.trim(), "0.5.0");
  const initialized = JSON.parse(expectSuccess([
    "init", ".", "--tools", "agents", "--json",
  ], project).stdout);
  assert.equal(initialized.created, 7);
  assert.ok(fs.existsSync(path.join(project, ".agents", "skills", "office-kit", "SKILL.md")));
  assert.equal(
    fs.existsSync(path.join(project, ".agents", "skills", "default-template-library")),
    false,
    "officekit init must not duplicate bundled templates into the project",
  );
  const updated = JSON.parse(expectSuccess([
    "update", ".", "--json",
  ], project).stdout);
  assert.equal(updated.unchanged, 7);

  const expectedTemplateCounts = new Map([
    ["document", 7],
    ["spreadsheet", 6],
    ["presentation", 7],
  ]);
  for (const [kind, expectedCount] of expectedTemplateCounts) {
    const search = JSON.parse(expectSuccess([
      "template", "search", "--kind", kind, "--max", "20", "--json",
    ], project).stdout);
    assert.equal(search.candidates.length, expectedCount, `${kind} bundled templates`);
    assert.equal(search.selectionMade, false);
    assert.deepEqual(search.invalid, []);
    assert.ok(search.searchedRoots.some((root) => root.source === "package-default"));
  }
  const none = JSON.parse(expectSuccess([
    "template", "search", "--kind", "presentation",
    "--purpose", "quantum entanglement laboratory protocol", "--json",
  ], project).stdout);
  assert.equal(none.retrievalStatus, "none");
  assert.deepEqual(none.candidates, []);
  const lazySearch = expectSuccess([
    "template", "search", "--kind", "document", "--max", "20", "--json",
  ], project, { NODE_DEBUG: "esm" });
  assert.doesNotMatch(
    lazySearch.stderr,
    /node_modules\/mupdf|src\/pdf\/mupdf|runtime\/office-kit\/main|src\/codecs\//iu,
    "template search must not initialize Office or PDF runtimes",
  );
  assert.equal(
    fs.existsSync(path.join(project, ".open-office-artifact-tool", "providers")),
    false,
    "init/search must not download a provider",
  );

  const packageMetadata = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
  const publicSpecifiers = Object.keys(packageMetadata.exports).map((subpath) =>
    subpath === "." ? packageMetadata.name : `${packageMetadata.name}/${subpath.slice(2)}`,
  );
  const taskPath = path.join(project, "four-formats.mjs");
  fs.writeFileSync(taskPath, [
    'import fs from "node:fs/promises";',
    'import {',
    '  DocumentFile, DocumentModel, PdfArtifact, PdfFile, Presentation,',
    '  PresentationFile, SpreadsheetFile, Workbook,',
    '} from "office-kit";',
    `const publicSpecifiers = ${JSON.stringify(publicSpecifiers)};`,
    "for (const specifier of publicSpecifiers) await import(specifier);",
    'const document = DocumentModel.create({ paragraphs: ["global CLI DOCX"] });',
    "const docx = await DocumentFile.exportDocx(document);",
    'await docx.save("global.docx");',
    "if ((await DocumentFile.importDocx(docx)).blocks[0].text !== 'global CLI DOCX') process.exit(11);",
    "const workbook = Workbook.create();",
    'workbook.worksheets.add("Data").getRange("A1:B2").values = [["Label", "Value"], ["global CLI XLSX", 7]];',
    "const xlsx = await SpreadsheetFile.exportXlsx(workbook);",
    'await xlsx.save("global.xlsx");',
    "if ((await SpreadsheetFile.importXlsx(xlsx)).worksheets.getItem('Data').getRange('B2').values[0][0] !== 7) process.exit(12);",
    "const presentation = Presentation.create();",
    'presentation.slides.add({ name: "Global CLI" }).shapes.add({',
    '  geometry: "textbox", text: "global CLI PPTX",',
    '  position: { left: 40, top: 40, width: 400, height: 80 },',
    "});",
    "const pptx = await PresentationFile.exportPptx(presentation);",
    'await pptx.save("global.pptx");',
    "if ((await PresentationFile.importPptx(pptx)).slides.count !== 1) process.exit(13);",
    'const pdf = await PdfFile.exportPdf(PdfArtifact.create({ pages: [{ text: "global CLI PDF" }] }));',
    'await pdf.save("global.pdf");',
    "if (!(await PdfFile.importPdf(pdf)).extractText().includes('global CLI PDF')) process.exit(14);",
    "for (const filename of ['global.docx', 'global.xlsx', 'global.pptx', 'global.pdf']) {",
    "  if ((await fs.stat(filename)).size < 100) process.exit(15);",
    "}",
    "console.log(JSON.stringify({",
    "  argv: process.argv.slice(2),",
    "  cwd: process.cwd(),",
    "  publicSubpaths: publicSpecifiers.length,",
    "}));",
    "",
  ].join("\n"));
  assert.equal(
    fs.existsSync(path.join(project, "node_modules")),
    false,
    "empty task project must start without node_modules",
  );
  const taskResult = JSON.parse(expectSuccess([
    "run", "four-formats.mjs", "--", "alpha", "two words",
  ], project).stdout);
  assert.deepEqual(taskResult.argv, ["alpha", "two words"]);
  assert.equal(taskResult.cwd, fs.realpathSync(project));
  assert.equal(taskResult.publicSubpaths, publicSpecifiers.length);
  assert.equal(
    fs.existsSync(path.join(project, "node_modules")),
    false,
    "officekit run must not install a project-local package",
  );

  const dependencyProject = path.join(temporary, "global-local-dependency-project");
  const dependencyRoot = path.join(dependencyProject, "node_modules", "local-probe");
  fs.mkdirSync(dependencyRoot, { recursive: true });
  fs.writeFileSync(
    path.join(dependencyRoot, "package.json"),
    `${JSON.stringify({
      name: "local-probe",
      version: "1.0.0",
      type: "module",
      exports: "./index.mjs",
    })}\n`,
  );
  fs.writeFileSync(path.join(dependencyRoot, "index.mjs"), "export default 41;\n");
  fs.writeFileSync(
    path.join(dependencyProject, "dependency-task.mjs"),
    'import value from "local-probe"; console.log(value + 1);\n',
  );
  assert.equal(
    expectSuccess(["run", "dependency-task.mjs"], dependencyProject).stdout.trim(),
    "42",
  );

  fs.writeFileSync(path.join(project, "exit-task.mjs"), "process.exitCode = 7;\n");
  const exitResult = execute(["run", "exit-task.mjs"], project);
  assert.equal(exitResult.status, 7, "officekit run must preserve task exit codes");
}
