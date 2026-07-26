# OfficeKit Skill

OfficeKit is the coordination entry point for broad or cross-format Office
work. It turns a request into an explicit artifact route, loads the required
Documents, Spreadsheets, Presentations, or PDF Skill, and decides whether zero
or one available template helps.

Install OfficeKit in the project where the Agent will work:

```sh
npm install github:w31r4/OfficeKit
npx officekit init
```

The initializer detects the Agent tools used by the project and installs the
OfficeKit entry point, Documents, Spreadsheets, Excel Live Control,
Presentations, PDF, and Template Creator in their project-local Skill
directories. Run `npx officekit update` after upgrading the package.

Repository templates and locally created `artifact-template-*` Skills are
queried through compact metadata and a local BM25F shortlist. The Agent decides
whether to select one, ask, or use none after it reviews the result and final
previews. Individual template Skills remain directly installable for explicit
invocation.

An uploaded DOCX, XLSX, or PPTX can be used once without registration. OfficeKit
preserves and inspects that file through the owning format Skill; Template
Creator is used only when the user explicitly wants a reusable local template.
