You are the author for a frozen presentation Skill experiment. Your arm is: kimi-concise. The lifecycle is 1→10. Do not compare arms or mention this arm in the artifact.

Read the arm instructions first: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/arms/kimi-concise/SKILL.md
Read the case file at /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-20260903/authors/management-report-10/kimi-concise/input/case.json
Use the OfficeKit repository at /Users/zfang/workspace/officekit-main-skill-eval-20260903; invoke its public CLI as: node /Users/zfang/workspace/officekit-main-skill-eval-20260903/bin/officekit.mjs ppj ...
Work only inside this trial workspace: /Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-20260903/authors/management-report-10/kimi-concise. Do not edit tracked repository files, do not use MJS/JSX as a presentation authoring surface, and do not run the full test suite.

Source PPTX (read-only; never overwrite): /Users/zfang/workspace/open-office-artifact-tool/tmp/reference-pptx-downloads/slidescarnival-business-infographic.pptx
Source SHA-256: 8db900eb9fbc5375d6b69eccffebd5ebb002f2f6641a89f19364a74e1d7e1e26
Target page: 7
Use officekit ppj import and perform exactly two serial edits: semantic first, then visual/delivery.


Follow the brief and acceptance contract exactly. Use only supplied facts; mark illustrative or assumed content. If a photo/icon is required, use the shared image route and record query, source, rights, hash, crop and alt text. Do not use host image generation.

Required outputs (even if a check fails, leave a precise report):
- outputs/deck.ppj
- outputs/deck.pptx when the codec is available
- outputs/previews/ (target page and adjacent page for 1→10; the page for 0→1)
- outputs/review.json or the CLI review output
- outputs/author-report.md explaining decisions, commands, failures and evidence type

Run the narrow sequence check → build → render → repair visible defects → review. For 1→10 also re-import the output and explain stable IDs, opaque content, source binding and non-target preservation. Never declare structural evidence to be PowerPoint playback evidence. Stop after the requested case; do not create a benchmark harness.