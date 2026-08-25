# What OfficeKit Means by a Presentation

OfficeKit treats a presentation as more than a sequence of styled pages. It is
a communication activity, an editable deck, a timed experience when presented,
and a native artifact that people may review, revise, reuse, or archive.

This definition guides how the Presentations Skill plans, composes, reviews,
and delivers PowerPoint work.

[简体中文](what-is-a-presentation.zh-CN.md)

## The audience outcome comes first

People rarely ask for slides because pages are the final goal. They want an
audience to know, understand, believe, decide, align, learn, report, act, or
retain something.

OfficeKit therefore begins with four questions:

1. Who is the audience?
2. What must change for them after the presentation?
3. What evidence supports that change?
4. How will the deck be delivered and used afterward?

The answers determine the narrative, density, visual language, notes, motion,
and native structure. Style does not substitute for an argument.

## A presentation has four connected forms

| Form | OfficeKit responsibility |
| --- | --- |
| Communication activity | Preserve the audience, purpose, expected outcome, evidence, and decision boundary. |
| Editable deck | Build an ordered document whose pages and objects can be reviewed, changed, rearranged, and reused. |
| Playback experience | Use pacing, transitions, animation, notes, and visual attention only when they support the delivery mode. |
| Native artifact | Keep a real PowerPoint object and package structure that can reopen, render, edit, and travel through later work. |

A visually polished screenshot is not enough if the file is difficult to edit,
loses inherited structure, or cannot survive another revision. A technically
valid file is not enough if it does not help its audience complete the intended
task.

## Delivery mode changes design

OfficeKit distinguishes three primary modes:

- **Live** — the speaker controls pace. Slides favor distance readability,
  strong focal points, deliberate reveals, and useful speaker notes.
- **Reader** — the deck must explain itself. Pages carry more context,
  traceable evidence, navigation, and fewer presentation-only effects.
- **Hybrid** — the deck serves both. The plan declares which mode leads and
  uses notes, appendices, or supporting pages for the other.

A user may also need the deck afterward as a decision record, handout,
reference, or reusable source. That after-use affects citations, notes,
editability, and how much context remains visible on the page.

If PowerPoint is a weak fit for the primary task, OfficeKit records that
limitation and continues when the user has requested a presentation. It does
not silently replace the deliverable with a document, dashboard, or website.

## Scenarios and design mechanisms are different

The scenario describes the organizational job. OfficeKit supports seven common
families: analysis and decision, business proposal, management report,
academic research, education and training, technical engineering, and brand
creative.

The design mechanism describes how information behaves: editorial restraint,
data review, technical architecture, visual narrative, academic evidence, or
brand launch. One scenario may use more than one mechanism. Neither layer is a
template or a fixed palette.

For every new deck, the Agent selects a primary scenario, identifies the design
authority, chooses a project-specific direction, and writes a concrete visual
grammar before composing pages.

## Design authority has a clear order

OfficeKit recognizes four design sources:

1. a user-supplied template or reference;
2. an explicit brand or design system;
3. a requested style transfer from observable evidence;
4. a self-directed design created for the task.

Authoritative source evidence wins. Scenario guidance may fill an unresolved
choice, but it must not overwrite a template's fonts, hierarchy, geometry,
assets, or other demonstrated rules. A template is useful because it captures
repeatable decisions; it is not required for good design.

## Quality has six layers

| Layer | Review question |
| --- | --- |
| Factual | Are claims, values, sources, and distinctions correct and traceable? |
| Communication | Are the audience, intended change, and requested action clear? |
| Narrative | Does the sequence create the intended understanding or decision? |
| Cognitive | Can the audience perceive, process, compare, and connect the information without unnecessary load? |
| Visual | Do hierarchy, typography, imagery, data graphics, spacing, and motion serve meaning? |
| Native and operational | Does the PowerPoint reopen, remain editable, preserve required structure, play reliably, and support later reuse? |

These layers require different evidence. Package inspection cannot prove facts.
Object counts cannot prove beauty. A compact text view cannot prove crops,
contrast, visual balance, or playback. OfficeKit keeps deterministic checks,
Agent judgment, visual review, and native-host evidence separate.

## What OfficeKit owns

The Agent decides what position to take and what message to communicate.
OfficeKit teaches the Agent how to turn that decision into a coherent native
presentation, checks deterministic constraints, preserves task state, and
makes limitations visible.

OfficeKit does not promise a universal aesthetic score, automatic factual
truth, a perfect one-shot final deck, or a replacement for human judgment.
Its deliverable is a strong, inspectable, editable working artifact that can
continue to improve through conversation.
