# OfficeKit Skill

OfficeKit is the coordination entry point for broad or cross-format Office
work. It turns a request into an explicit artifact route, loads the required
Documents, Spreadsheets, Presentations, or PDF Skill, and decides whether zero
or one available template helps.

Install OfficeKit once, then initialize a project:

```sh
curl -fsSL https://github.com/w31r4/OfficeKit/releases/latest/download/install.sh | sh
```

On Windows PowerShell:

```powershell
irm https://github.com/w31r4/OfficeKit/releases/latest/download/install.ps1 | iex
```

Open a new terminal, then initialize the project:

```sh
officekit init
```

The self-contained macOS arm64, Linux x64, and Windows x64 builds carry Node
24.18.0, OfficeKit, its Skills, and the default templates.

The initializer detects the Agent tools used by the project and installs the
OfficeKit entry point, Documents, Spreadsheets, Excel Live Control,
Presentations, PDF, and Template Creator in their project-local Skill
directories. Run `officekit update` after installing a newer OfficeKit release.
Skill tasks can use `officekit run task.mjs`; multi-step artifact work can list
workspace-local tasks, open or create one task REPL, stage inputs, review and
commit stable revisions, then publish. The project does not need a local
`office-kit` dependency. See the installed OfficeKit Skill's
`references/repl.md` for the JSONL and recovery contract.

The twenty bundled templates, project templates, and locally created
`artifact-template-*` Skills are queried with
`officekit template search ... --json`. Compact English metadata feeds a local
BM25F shortlist. The Agent decides whether to select one, ask, or use none after
it reviews the result and final previews. Individual template Skills remain
directly invokable.

An uploaded DOCX, XLSX, or PPTX can be used once without registration. OfficeKit
preserves and inspects that file through the owning format Skill; Template
Creator is used only when the user explicitly wants a reusable local template.
