# Legacy Presentation implementations

This directory contains retired Presentation authoring material kept for
history, comparison, and controlled migration work. It is not a public
OfficeKit authoring surface.

`src/`, `scripts/`, and `test/` contain the retired JavaScript object model,
wire mapper, PowerPoint REPL facade, review helpers, benchmarks, and tests.
They are outside the npm package and active test graph.

`mjs/office-artifact-tool` is the pinned public reference submodule at
`73c99c67ca7bbaa82cec0b158c647db583dcd970`. It contains the former MJS
Presentation Skill, executable Grid layout modules, old template workflows,
and the wider upstream comparison tree.

Current Presentation creation uses PPJ and current Skills under `skills/`.
Do not add new template behavior here or import this legacy runtime into the
OfficeKit package.
