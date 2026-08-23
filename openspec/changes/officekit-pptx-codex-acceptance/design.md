# Design

The platform has four one-way layers:

1. Immutable JSON task definitions bind source hashes, public targets, desired
   values, target pages, and evaluator-owned allowed package footprints.
2. A deterministic runner imports the packed public package from clean source
   bytes, performs one typed edit, exports, and second-imports.
3. An independent oracle compares source and output packages, relationships,
   masked target XML/SVG, nested packages, and native-rendered page pixels.
4. A Codex harness creates an isolated clean-install workspace per trial,
   installs only the candidate tarball and its Skill, runs one ephemeral Agent
   context, and evaluates the resulting durable task plus output with layer 3.

The oracle never consumes an OfficeKit edit plan as authority. Runtime
metadata may be retained as diagnostics, but a mismatch between metadata and
source/output bytes is a failure. Pixel rendering may cache by content hash;
cache reuse is valid only after byte identity has already been proven.

Every output and evidence path is create-only. Source files are copied into an
isolated read-only input directory and re-hashed after execution. Agent traces
and authored scripts are scanned for forbidden implementation paths. A failed
Agent, missing external renderer, unsupported public operation, or oracle
mismatch is recorded without retrying through another engine.
