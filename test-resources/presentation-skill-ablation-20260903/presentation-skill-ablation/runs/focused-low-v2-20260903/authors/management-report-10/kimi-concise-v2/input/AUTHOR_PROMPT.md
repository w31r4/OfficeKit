You are authoring one frozen OfficeKit presentation edit experiment. This is the kimi-concise-v2 arm; do not mention arms in the artifact. The project is experimental, not production.

Read the arm instructions first: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/arms/kimi-concise/SKILL.md
Read the case file: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/management-report-10/kimi-concise-v2/input/case.json
Use the OfficeKit public CLI at: node /Users/zfang/workspace/officekit-main-skill-eval-20260903/bin/officekit.mjs ppj ...
Work only in this workspace: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/management-report-10/kimi-concise-v2

The source PPTX has already been projected to outputs/deck.ppj and its source/media assets are already beside it. Do NOT re-import the source, and do not overwrite the input PPTX. Start by inspecting the assigned target page in the existing PPJ.

Follow the case brief and do exactly two serial stages: semantic edit first, then visual/delivery edit. If a requested surface does not exist on the assigned page, fail closed and leave that stage unchanged; do not move the edit to another page or invent data. If the surface exists, make the smallest source-bound edit that satisfies the brief. Preserve stable IDs, opaque content, source binding, and non-target pages/parts.

Use only supplied facts; mark illustrative/pending text when the case does not supply a value. Do not use host image generation or write a benchmark harness. Use PPJ directly, not MJS/JSX.

After editing run: ppj check outputs/deck.ppj --json; ppj build outputs/deck.ppj -o outputs/deck.pptx --json; ppj render outputs/deck.ppj -o outputs/previews --pages <target and adjacent pages> --json; inspect the rendered PNGs and repair the responsible layer if needed; ppj review outputs/deck.ppj --json; and ppj import outputs/deck.pptx -o evidence/reimport.ppj --json. Do not cover an error with a new shape. Record commands, decisions, hard gates, source hash, stable IDs, and any limitation in outputs/author-report.md. Leave outputs/deck.ppj, outputs/deck.pptx when possible, outputs/previews/, outputs/review.json or equivalent, and the report even if a stage is blocked.

Case brief:
Import the assigned source page. Step 1: change the requested KPI text and one table value using the existing native fields. Step 2: improve the page's hierarchy with a local typography/spacing adjustment and, only if needed, a crop or layer adjustment; preserve the source's communication contract and all non-target content.

Target/source metadata:
- fixture: business-infographic
- source path: /Users/zfang/workspace/open-office-artifact-tool/tmp/reference-pptx-downloads/slidescarnival-business-infographic.pptx
- source SHA-256: 8db900eb9fbc5375d6b69eccffebd5ebb002f2f6641a89f19364a74e1d7e1e26
- target page contract: 7
- edit steps: native KPI/table value edit -> typography/spacing or crop repair
