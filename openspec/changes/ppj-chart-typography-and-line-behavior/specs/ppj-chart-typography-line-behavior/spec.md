## ADDED Requirements

### Requirement: Bounded chart-title text style
PPJ SHALL allow an inline or named chart style to declare a chart-title text style whose initial compiler-owned property is explicit font size, and NativeAOT SHALL preserve that state through authored PPTX build and recognized PPJ reimport.

#### Scenario: Authored title size round trips
- **WHEN** a PPJ chart with a non-empty title declares `titleTextStyle.fontSize`
- **THEN** the native PPTX contains the bounded title size and a subsequent projection restores the same PPJ value

#### Scenario: Title style without title is rejected
- **WHEN** a PPJ chart declares title text style without a non-empty title
- **THEN** validation or compilation fails before native output changes

### Requirement: Canonical line-chart behavior
PPJ SHALL allow a line chart style to declare explicit smooth interpolation and direct color variation, SHALL preserve explicit smooth false separately from omission, and SHALL compile the state through the existing canonical ChartSpace owner.

#### Scenario: Smooth line behavior round trips
- **WHEN** a PPJ line chart declares `smooth: true` or `smooth: false`
- **THEN** authored build and recognized reimport preserve that explicit value

#### Scenario: Direct color variation round trips
- **WHEN** a PPJ line chart declares `varyColors: true`
- **THEN** authored build emits the canonical native state and recognized reimport restores `varyColors: true`

#### Scenario: Incompatible chart rejects line behavior
- **WHEN** a non-line PPJ chart declares `smooth` or `varyColors`
- **THEN** compilation fails with a path-specific unsupported-state diagnostic before writing output

### Requirement: Imported style mutation remains capability-bound
Recognized source charts SHALL project title and line behavior for inspection, but changing that state SHALL fail closed unless the source projection issues a dedicated capability for the field.

#### Scenario: Unsupported source style mutation is rejected
- **WHEN** an Agent changes projected title typography or line behavior on a source-bound chart that only issues title, data, and frame capabilities
- **THEN** build rejects the style mutation and preserves the source package unchanged
