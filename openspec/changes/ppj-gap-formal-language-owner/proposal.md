# Proposal: make the formal text language owner explicit

Extend the existing bounded text-owner precedence profile with the already
modeled `text.language` field. A declared precedence rule will let authored
PPJ choose the BCP-47 language tag from run, paragraph, element, named style,
layout, master, theme, or default sources before lowering it to the direct
DrawingML `a:rPr/@lang` owner.

This closes one concrete language-owner gap without inferring fonts, shaping
complex scripts, evaluating host proofing behavior, or editing imported theme
XML. Source-bound inherited language graphs remain read-only; the existing
direct `fontLanguage` leaf remains the only imported edit surface.
