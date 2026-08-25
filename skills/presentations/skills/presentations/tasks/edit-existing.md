# Edit an existing PPTX

Use this route for a local PowerPoint file that remains the design and content
authority.

## 1. Stage, import, and inspect

Stage the source with `ctx.input`, import the managed copy, and run targeted
`presentation.inspect()` calls. Resolve the requested page and object before
choosing an edit primitive.

Classify the target as typed-editable, native-leaf-editable,
source-derived-reusable, or opaque-preserved. Load
[advanced imported editing](../references/advanced-imported-editing.md) only
for the relevant advanced object.

## 2. Declare scope

Use an `edit-existing` authoring plan for broad or resumable work. Preserve the
source design grammar and editorial voice. List only the pages that the user
asked to change; a global redesign requires explicit scope.

For copy changes, load the sibling
[`presentation-editorial-trim`](../../presentation-editorial-trim/SKILL.md)
Skill. Record exact facts, citations, protected wording, and target page IDs
before editing. Inherit the deck's existing title rhythm and terminology; do
not normalize non-target pages to the new wording.

## 3. Apply a bounded edit

Prefer typed APIs. Use `editNativeLeaf` only with inspect-issued IDs and the
current expected hash. Use source component/slide reuse only with its inspected
capability. Avoid raw XML, XPath, relationship IDs, and whole-package rebuilds.

Export after a source-bound edit. Reimport and reinspect before the next edit;
capability IDs do not survive revision changes.

If copy reflows, render the changed pages and run the editorial page-fit pass.
Do not alter layout or neighboring objects when the requested scope is
copy-only unless fit cannot be restored safely; report that limitation instead.

## 4. Review the edit boundary

Call `reviewArtifact` with the source or latest reviewed revision as
`baseline`, the active `authoringPlan`, and exact `changedPageIds`. Treat an
undeclared page change as a failed review. Render affected pages and compare
non-target evidence before committing.

Follow [Review and deliver](review-deliver.md).
