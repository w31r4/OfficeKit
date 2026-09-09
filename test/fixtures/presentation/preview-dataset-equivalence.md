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

This is authored dataset lowering, not a source-bound workbook edit, external
Office comparison, human calibration or proof for every encoding channel.
Failures of either side remain fatal after independent regressions run.
