# Why OfficeKit Uses PPJ

PPJ (PPT Program JSON) is OfficeKit's declarative source format for
presentations. A `.ppj` file is strict UTF-8 JSON. It describes the current
deck: its communication intent, design grammar, pages, ordered elements,
assets, motion, and source bindings.

This document explains why OfficeKit separates a general-purpose Agent from a
finite presentation language.

[简体中文](why-ppj.zh-CN.md) · [What OfficeKit Means by a Presentation](what-is-a-presentation.md)

## The artifact and the work are different things

Creating a presentation may require open-ended work: researching evidence,
calculating values, comparing alternatives, finding images, and revising after
visual review. A general-purpose Agent or programming language is useful for
that work.

The resulting presentation is different. At any revision it has a finite page
count, finite objects, explicit text, assets, geometry, stacking order, and
relationships. OfficeKit needs to validate, diff, compile, reopen, review, and
resume that state. PPJ represents this finite artifact directly.

```text
Agent or optional script
  research · reason · calculate · generate
                    ↓
                  PPJ
  finite state · typed objects · stable identity
                    ↓
           NativeAOT C# compiler
                    ↓
                  PPTX
```

The Agent may use a Turing-complete language around PPJ. PPJ itself does not
need arbitrary computation.

## What Turing completeness means here

A Turing-complete language can express any computable procedure, given enough
time and memory. JavaScript can branch, loop, recurse, create data dynamically,
and interact with external systems. Those powers are valuable when the route
to an answer is not known in advance.

They also make arbitrary programs impossible to analyze perfectly. The halting
problem gives the classic boundary. Assume a perfect function exists:

```text
willStop(program, input)
```

It claims to decide whether any program will terminate. Now construct a
program `contrary` that asks what happens when it receives itself as both the
program and the input. If `willStop` predicts termination, `contrary` loops; if
it predicts a loop, `contrary` terminates. Either answer contradicts the
prediction. A universal perfect `willStop` cannot exist.

This does not make JavaScript unsafe or unsuitable. Real systems use timeouts,
budgets, cancellation, and restricted capabilities. The lesson is narrower:
OfficeKit cannot accept an arbitrary program and always prove its termination,
mutation footprint, deterministic output, or safe resumability before running
it.

## Why a presentation DSL is a better persistent source

PPJ admits only bounded presentation constructs. It has typed elements,
stable IDs, ordered arrays, finite component expansion, local asset references,
and explicit source capabilities. It has no functions, recursion, `while`,
network access, raw OOXML, or arbitrary expressions.

That boundary gives OfficeKit useful guarantees before touching a PPTX:

- the program can be parsed and validated completely;
- component expansion has a known upper bound;
- changed nodes and affected package parts can be calculated;
- the same validated input can compile deterministically;
- a fresh Agent can reopen the file without restoring a JavaScript heap;
- unsupported imported edits can fail before source content is damaged;
- program revisions can be diffed, reviewed, and resumed as ordinary data.

JSON alone is only a data syntax. PPJ becomes a domain-specific language when
OfficeKit defines its schema, semantics, validation, compiler, and review
contract.

## Browser and computer control need a different split

A browser or desktop task operates in an open, changing environment. A login
dialog, redirect, timeout, modal, or unexpected page may require a new branch,
retry, or recovery strategy. A general controller is useful there:

```text
observe → decide → act → observe again
```

The individual actions should still be bounded and typed, such as `navigate`,
`click`, `type`, `scroll`, `wait`, and `screenshot`. The controller may be
Turing-complete; the action protocol remains auditable.

Presentations use the same separation at a different boundary:

| Layer | Appropriate form |
| --- | --- |
| Research, planning, calculation, iteration | General Agent or optional script |
| Persistent deck state | PPJ |
| Imported local mutation | Typed, source-bound edit plan |
| Native file generation | Deterministic C# compiler |
| Observation and correction | Render and review evidence |

The rule is not “DSL everywhere.” Open-world exploration benefits from a
general controller. A finite artifact benefits from a declarative source.

## Authored and imported decks have different authority

For an OfficeKit-authored deck, PPJ is the source program and PPTX is the
compiled artifact.

For an arbitrary third-party PPTX, the original package remains the source of
unknown native content. OfficeKit projects understood objects into typed PPJ
elements and exposes bounded imported capabilities through native references.
Unknown graphs remain opaque in the source package. Editing the PPJ produces a
verified local edit plan rather than reconstructing the whole file.

This asymmetry is deliberate. It lets an Agent use one readable programming
surface without claiming that every OOXML feature has been fully decompiled.

## The intended OfficeKit workflow

```text
create:  deck.ppj → check → build → render → review → revise
import:  deck.pptx → project PPJ → edit → verified local build
resume:  reopen last valid PPJ revision → inspect → continue
```

PPJ is therefore not a weaker replacement for JavaScript. It is the stable
contract between an intelligent, adaptable Agent and a deterministic native
presentation compiler.
