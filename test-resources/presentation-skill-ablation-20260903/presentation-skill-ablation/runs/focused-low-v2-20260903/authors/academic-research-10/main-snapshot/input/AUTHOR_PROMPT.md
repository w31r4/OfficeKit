You are authoring one frozen OfficeKit presentation edit experiment. This is the main-snapshot arm; do not mention arms in the artifact. The project is experimental, not production.

Read the arm instructions first: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/arms/current-production/SKILL.md
Read the case file: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/academic-research-10/main-snapshot/input/case.json
Use the OfficeKit public CLI at: node /Users/zfang/workspace/officekit-main-skill-eval-20260903/bin/officekit.mjs ppj ...
Work only in this workspace: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/academic-research-10/main-snapshot

The source PPTX has already been projected to outputs/deck.ppj and its source/media assets are already beside it. Do NOT re-import the source, and do not overwrite the input PPTX. Start by inspecting the assigned target page in the existing PPJ.

Follow the case brief and do exactly two serial stages: semantic edit first, then visual/delivery edit. If a requested surface does not exist on the assigned page, fail closed and leave that stage unchanged; do not move the edit to another page or invent data. If the surface exists, make the smallest source-bound edit that satisfies the brief. Preserve stable IDs, opaque content, source binding, and non-target pages/parts.

Use only supplied facts; mark illustrative/pending text when the case does not supply a value. Do not use host image generation or write a benchmark harness. Use PPJ directly, not MJS/JSX.

After editing run: ppj check outputs/deck.ppj --json; ppj build outputs/deck.ppj -o outputs/deck.pptx --json; ppj render outputs/deck.ppj -o outputs/previews --pages <target and adjacent pages> --json; inspect the rendered PNGs and repair the responsible layer if needed; ppj review outputs/deck.ppj --json; and ppj import outputs/deck.pptx -o evidence/reimport.ppj --json. Do not cover an error with a new shape. Record commands, decisions, hard gates, source hash, stable IDs, and any limitation in outputs/author-report.md. Leave outputs/deck.ppj, outputs/deck.pptx when possible, outputs/previews/, outputs/review.json or equivalent, and the report even if a stage is blocked.

Case brief:
Import the assigned research page. Step 1: update the supplied sample-size note and one result label, retaining uncertainty and source attribution. Step 2: repair the local table/chart alignment and label spacing so no number, interval, or footnote is hidden; do not change the conclusion or unrelated source parts.

Target/source metadata:
- fixture: nasa-mms
- source path: /Users/zfang/workspace/open-office-artifact-tool/tmp/reference-pptx-downloads/nasa-mms-machine-learning.pptx
- source SHA-256: 531c82797fde09b1ebe1e868ca9cd44c3e2f675dc8f09f58b54bab6a62629723
- target page contract: 9
- edit steps: native note/result-label edit -> table/chart alignment and label repair
