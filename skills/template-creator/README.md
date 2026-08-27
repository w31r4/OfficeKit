# Template Creator

Create or update source-backed local templates from a `.docx` or `.xlsx`
reference and one representative PNG preview. The creator validates the Office
package, preserves the reference bytes, records hashes, and publishes the Skill
atomically below `${OFFICE_KIT_HOME:-~/.office-kit}/skills`.

PowerPoint templates use the separate `presentation-template-creator`. A PPT
template is style guidance plus original visual examples; it does not retain a
PPTX or fixed page layout.

```sh
officekit run skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs -- \
  --reference-path /absolute/path/reference.docx \
  --preview-path /absolute/path/preview.png \
  --display-name "Decision memo" \
  --description "Draft a concise decision memorandum with evidence and recommendations."
```

The generated DOCX/XLSX template uses schema v2 and contains `SKILL.md`,
`artifact-template.json`, `agents/agent.yaml`, a retained
`assets/reference.<ext>`, and `assets/preview.png`.
