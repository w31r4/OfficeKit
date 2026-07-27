# OfficeKit Skill

OfficeKit is the coordination entry point for broad or cross-format Office
work. It turns a request into an explicit artifact route, loads the required
Documents, Spreadsheets, Presentations, or PDF Skill, and decides whether zero
or one available template helps.

Install OfficeKit once, then initialize a project:

```sh
npm install -g github:w31r4/OfficeKit
officekit init
```

The initializer detects the Agent tools used by the project and installs the
OfficeKit entry point, Documents, Spreadsheets, Excel Live Control,
Presentations, PDF, and Template Creator in their project-local Skill
directories. Run `officekit update` after upgrading the global package. Skill
tasks use `officekit run task.mjs`, so the project does not need a local
`office-kit` dependency.

The twenty bundled templates, project templates, and locally created
`artifact-template-*` Skills are queried with
`officekit template search ... --json`. Compact English metadata feeds a local
BM25F shortlist. The Agent decides whether to select one, ask, or use none after
it reviews the result and final previews. Individual template Skills remain
directly invokable.

An uploaded DOCX, XLSX, or PPTX can be used once without registration. OfficeKit
preserves and inspects that file through the owning format Skill; Template
Creator is used only when the user explicitly wants a reusable local template.
