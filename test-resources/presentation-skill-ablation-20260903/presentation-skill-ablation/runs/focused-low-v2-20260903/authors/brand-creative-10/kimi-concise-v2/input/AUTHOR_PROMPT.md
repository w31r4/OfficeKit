You are authoring one frozen OfficeKit presentation edit experiment. This is the kimi-concise-v2 arm; do not mention arms in the artifact. The project is experimental, not production.

Read the arm instructions first: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/arms/kimi-concise/SKILL.md
Read the case file: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/brand-creative-10/kimi-concise-v2/input/case.json
Use the OfficeKit public CLI at: node /Users/zfang/workspace/officekit-main-skill-eval-20260903/bin/officekit.mjs ppj ...
Work only in this workspace: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903/authors/brand-creative-10/kimi-concise-v2

The source PPTX has already been projected to outputs/deck.ppj and its source/media assets are already beside it. Do NOT re-import the source, and do not overwrite the input PPTX. Start by inspecting the assigned target page in the existing PPJ.

Follow the case brief and do exactly two serial stages: semantic edit first, then visual/delivery edit. If a requested surface does not exist on the assigned page, fail closed and leave that stage unchanged; do not move the edit to another page or invent data. If the surface exists, make the smallest source-bound edit that satisfies the brief. Preserve stable IDs, opaque content, source binding, and non-target pages/parts.

Use only supplied facts; mark illustrative/pending text when the case does not supply a value. Do not use host image generation or write a benchmark harness. Use PPJ directly, not MJS/JSX.

After editing run: ppj check outputs/deck.ppj --json; ppj build outputs/deck.ppj -o outputs/deck.pptx --json; ppj render outputs/deck.ppj -o outputs/previews --pages <target and adjacent pages> --json; inspect the rendered PNGs and repair the responsible layer if needed; ppj review outputs/deck.ppj --json; and ppj import outputs/deck.pptx -o evidence/reimport.ppj --json. Do not cover an error with a new shape. Record commands, decisions, hard gates, source hash, stable IDs, and any limitation in outputs/author-report.md. Leave outputs/deck.ppj, outputs/deck.pptx when possible, outputs/previews/, outputs/review.json or equivalent, and the report even if a stage is blocked.

Case brief:
Import the assigned launch page. Step 1: replace the supplied product promise text with the approved copy and update one specification label. Step 2: replace or crop the hero image and repair the foreground contrast/type layer so the reveal is premium, legible, editable, and unchanged outside the requested page.

Target/source metadata:
- fixture: professional-minimalist
- source path: /Users/zfang/workspace/open-office-artifact-tool/tmp/reference-pptx-downloads/slidescarnival-professional-minimalist.pptx
- source SHA-256: ed26f074e25361e4baf76d1cc2508596655b9d4f3fc2c659ecd962f76e0547ab
- target page contract: 4
- edit steps: approved promise/specification edit -> hero image/crop and foreground contrast repair
