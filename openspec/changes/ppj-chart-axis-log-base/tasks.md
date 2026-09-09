## 1. Field implementation

- [x] 1.1 Add optional logBase across schema, wire, native axis codec and PPJ compile/project/edit paths; verify focused authored and source-bound round trips.
- [x] 1.2 Protect invalid and ambiguous logarithmic axes; verify rejected input and preserved native chart bytes.
- [x] 1.3 Preserve imported XLSX logBase through JavaScript chart edits now that the shared native codec recognizes it; verify title edits retain the wire field.

## 2. Evidence and documentation

- [x] 2.1 Add minimal experiments for regular, numeric x, radar and secondary value axes, token edits and deletion; run focused codec regressions.
- [x] 2.2 Update PPJ reference, capability registry and F-07 backlog; regenerate protobuf bindings and verify proto:check and strict OpenSpec validation.
