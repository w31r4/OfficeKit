# Dataset / explicit chart preview pair

`preview-dataset-equivalence.json` is a fixed authored-input pair consumed by
`test/ppj-preview-scene-native.mjs`, using `examples/ppj/minimum.ppj` as its shared
document base. The explicit categories and series are written independently;
the test never uses compiler output to construct the expected input.

The dataset interleaves Alpha and Beta rows, mixes array and object rows, and
selects channels by both name and index. Expected series are Alpha `[1,null,0]`
and Beta `[0,4,5]`, in that order. The explicit series IDs follow the compiler's
documented-in-code `series-{index+1}` generation in `PpjProgramModels.CanonicalDataset`.

The actual NativeAOT test asserts names, frame, ordered values and missing indexes,
two isolated Alpha observations and exactly one continuous Beta line segment.
The entire typed chart payload and raw page raster must equal the explicit side.
Both sides must retain the original chart owner; scene-on/off compilation must
produce identical candidate bytes for each input, and the inputs stay unchanged.
Fixture bytes participate in the report's before/after identity check.

The same fixed data pair also runs as a vector heatmap, in addition to native
line. Its style is fixed in the test: black-to-red linear domain [0,5], green
missing fill, no color bar/value labels and zero cell gap. The compiler must
produce a group with six cells and separate generated scene addresses, all
owned by the original chart. Ordered colors must be 330000/00FF00/000000 and
000000/CC0000/FF0000. Fixed expected row/column geometry and all six center
pixels are asserted, distinguishing the missing cell from true zero. The full
typed group (including generated children) and raw raster must equal the
explicit-data side. There is no JS heatmap layout or color interpolation.

This is authored dataset lowering, not a source-bound workbook edit, external
Office comparison, human calibration or proof for every encoding channel.
Failures of either side remain fatal after independent regressions run.
