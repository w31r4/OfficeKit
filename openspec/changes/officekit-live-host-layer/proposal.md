# OfficeKit Live Host Layer

## Why

Excel Live already has the local pairing, HTTPS, queue, idempotency, and audit
transport required for an open Office document. PowerPoint should use the same
transport instead of creating a second bridge, while keeping host execution
typed and isolated.

## Scope

- expose the shared Live transport and protocol 1 adapter boundary;
- preserve `officekit excel ...` for compatibility;
- add a PowerPoint add-in and `officekit live ... --app powerpoint` commands;
- add a lazy `ctx.powerpoint` REPL facade and a dedicated live Skill;
- make Windows x64 the first real desktop acceptance platform.

Word Live, raw Office.js, cloud relays, silent closed-file fallback, and Office
wire changes are out of scope.
