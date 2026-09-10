# Proposal: add source-bound static text-field type lifecycle

The PPJ projection already exposes `text.paragraphs[].runs[].field.type`, but
source-bound compilation rejects changing it. Add a narrow lifecycle for this
field so an imported static DrawingML field can change its `a:fld/@type` token
without rebuilding the text body.

The profile is deliberately bounded. It issues a leaf only for a valid static
field, requires the existing field ID, cached display text, automatic marker,
run topology, and field topology to stay unchanged, and patches one direct XML
attribute in the owning slide. Automatic field evaluation, field ID changes,
cached display refresh, and unknown field descendants remain source-owned.
