# How: Construction Contract

After the communication and route decisions, write a deck-specific Design
Grammar and page attention contract:

claim → relationship → carrier → reading order → protected evidence
→ occupancy → layer order → type/shape/image treatment

Read only the implementation references required by the carrier:

- references/ppj.md for the strict JSON language and limits;
- references/fonts.md and references/text.md for hierarchy and type;
- references/shapes.md for geometry, lines, and connectors;
- references/charts-and-tables.md for data marks and tables;
- references/image-sourcing.md and references/media-and-layers.md for
  images, masks, backgrounds, opacity, crop, and z-order;
- references/motion.md only when delivery or the brief authorizes motion;
- references/imported-native-ref.md for source-bound edits;
- references/components-and-templates.md for a selected style authority.

Compose every page in PPJ. Arrays are semantic: page order is pages[], and
element order is true back-to-front z-order. Review the rendered page before
adding decoration or motion.
