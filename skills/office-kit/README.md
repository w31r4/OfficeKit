# OfficeKit Skill

OfficeKit is the coordination entry point for broad or cross-format Office
work. It turns a request into an explicit artifact route, loads the required
Documents, Spreadsheets, Presentations, or PDF Skill, and decides whether zero
or one available template helps.

Install the coordinated core Skills together:

```sh
npx skills add w31r4/office-kit \
  --skill office-kit documents spreadsheets excel-live-control presentations pdf template-creator \
  --yes
```

The installer only deploys Skill instructions and resources. OfficeKit does
not replace the format-specific workflows or provide a second file codec.

Repository templates and locally created `artifact-template-*` Skills are
queried through compact metadata and a local BM25F shortlist. The Agent decides
whether to select one, ask, or use none after it reviews the result and final
previews. Individual template Skills remain directly installable for explicit
invocation.

An uploaded DOCX, XLSX, or PPTX can be used once without registration. OfficeKit
preserves and inspects that file through the owning format Skill; Template
Creator is used only when the user explicitly wants a reusable local template.
