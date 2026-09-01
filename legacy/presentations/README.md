# Legacy Presentation implementations

This directory contains the frozen reference used by controlled Presentation
comparisons. It is not a public OfficeKit authoring surface.

`mjs/office-artifact-tool` is the pinned public reference submodule at
`73c99c67ca7bbaa82cec0b158c647db583dcd970`. It contains the former MJS
Presentation Skill, executable Grid layout modules, old template workflows,
and the wider upstream comparison tree.

The retired duplicate JavaScript runtime, scripts, benchmarks, and tests were
removed after PPJ became the sole Presentation product boundary. Historical
source remains available from Git rather than as a second implementation tree.

Current Presentation creation uses PPJ and current Skills under `skills/`.
Do not add new template behavior here or import this legacy runtime into the
OfficeKit package.
