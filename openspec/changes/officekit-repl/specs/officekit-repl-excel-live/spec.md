## ADDED Requirements

### Requirement: Lazy Excel Live access from a REPL

The REPL SHALL expose a lazy `ctx.excel` facade for the existing local Excel
bridge. The facade SHALL provide typed `doctor`, `sessions`, `execute`, and
`disconnect` operations and SHALL reuse the existing Excel request protocol;
it SHALL NOT expose arbitrary Office.js evaluation.

#### Scenario: Use an existing Excel session

- **WHEN** code calls `await ctx.excel.sessions()` after the user has connected
  the OfficeKit add-in in Excel
- **THEN** the facade returns the existing session descriptors without creating
  a second bridge protocol or object model

#### Scenario: Execute a typed operation

- **WHEN** code calls `ctx.excel.execute(request)` with a schema-valid request
- **THEN** the facade forwards it to the existing bridge and returns the
  protocol's result, audit data, and `maybeApplied` state unchanged

### Requirement: Explicit Excel setup boundary

Accessing `ctx.excel` or starting a REPL SHALL NOT install certificates, trust
the local certificate, download an add-in, or start a bridge with no explicit
Excel operation. Installation and uninstall remain explicit `officekit excel`
control-plane commands.

#### Scenario: Start without Excel setup

- **WHEN** the Agent starts a REPL on a machine where Excel Live is not
  installed and performs only file work
- **THEN** no Excel state directory, certificate, bridge process, or network
  request is created

#### Scenario: Report an unavailable live session

- **WHEN** code calls `ctx.excel.execute(request)` without a connected workbook
- **THEN** the facade returns the existing typed unavailable-session error with
  a retryability indicator and does not silently switch to XLSX file editing

### Requirement: Excel Live Skill routing

The Excel Live Control Skill SHALL teach the Agent to launch the REPL, run
`doctor` and `sessions` before mutations, use typed operation requests, reread
the target range after uncertain execution, and call `disconnect` explicitly.
It SHALL preserve the existing explicit-install and desktop-platform limits.

#### Scenario: Route an open workbook task

- **WHEN** the user asks to change a workbook that is currently open in Excel
- **THEN** the Skill selects `ctx.excel` rather than the ordinary XLSX file
  workflow and requires a read-back verification before reporting completion

#### Scenario: Route a closed XLSX task

- **WHEN** the user asks to create or edit a closed `.xlsx` file
- **THEN** the Skill uses the Spreadsheet API in the REPL and does not require an
  Excel Live session
