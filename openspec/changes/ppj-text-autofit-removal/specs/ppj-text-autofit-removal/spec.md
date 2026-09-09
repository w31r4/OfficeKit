## Purpose

Make direct text AutoFit choices and shrink-text percentages independently removable and restorable in source-bound PPJ edits.

## ADDED Requirements

### Requirement: AutoFit choice presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL distinguish explicit none, shrink-text, resize-shape and absence. Removing autoFit and its dependent profile SHALL remove the native choice and preserve surrounding state; all modes SHALL be restorable.

#### Scenario: Delete and restore a populated shrink-text choice
- **WHEN** autoFit and normalAutoFit are removed from projected source text and a mode is restored
- **THEN** absence produces no AutoFit child, explicit none produces noAutofit, and fresh projection matches while non-target XML/ZIP state is preserved

### Requirement: Independent normal AutoFit percentages

With shrink-text retained, deleting one normalAutoFit field SHALL remove only its direct percentage attribute. Removing the profile SHALL remove both direct percentages while retaining shrink-text. Existing ranges, precision and dependency validation SHALL remain enforced.

#### Scenario: Remove and restore percentages
- **WHEN** fontScale, lineSpacingReduction, or the whole profile is removed and restored
- **THEN** native presence and values match fresh projection, explicit zero lineSpacingReduction remains present, and the other field is retained for single-field operations

### Requirement: Guarded style removal and preview honesty

Whole styles containing only supported removable fields including autoFit/normalAutoFit SHALL support removal and compact table restoration. Missing authority, other fields and noncanonical source topology SHALL retain their existing rejection boundaries. Preview SHALL keep text-reflow limitations visible.

#### Scenario: Delete a simple AutoFit style
- **WHEN** an authorized simple style is removed and then restored
- **THEN** direct choice/profile state disappears and returns without rewriting text topology or claiming host reflow acceptance
