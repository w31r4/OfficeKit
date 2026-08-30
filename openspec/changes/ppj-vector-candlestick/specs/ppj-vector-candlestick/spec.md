## ADDED Requirements

### Requirement: PPJ SHALL express one bounded candlestick series

The language SHALL accept `chartType: "candlestick"` with ordered string
categories, close `values`, required `highValues` and `lowValues`, and optional
`openValues`. It SHALL reject channel-length mismatch, missing/non-finite values,
more than 64 observations, multiple series, invalid price inequalities and
ordinary series fields outside the profile.

#### Scenario: Valid OHLC program

- **WHEN** all four channels align and every open/close lies between low/high
- **THEN** validation succeeds before native compilation

#### Scenario: Invalid financial evidence

- **WHEN** a low exceeds a high or an open/close lies outside the range
- **THEN** validation rejects the PPJ without output

### Requirement: NativeAOT SHALL compile editable vector candlesticks

The compiler SHALL emit one deterministic native group containing editable
wicks, OHLC bodies or HLC close ticks, bounded axes and labels. It SHALL NOT add
a raster image or claim a ChartPart.

#### Scenario: OHLC build

- **WHEN** `openValues` is present
- **THEN** each observation has one wick and one rise/fall body

#### Scenario: HLC build

- **WHEN** `openValues` is omitted
- **THEN** each observation has one wick and one close tick without a body

### Requirement: Recovery SHALL distinguish semantic and native state

An authored PPTX with its embedded PPJ SHALL restore the exact candlestick
program. If the embedded program is removed, projection SHALL expose an
ordinary editable group and SHALL NOT infer candlestick semantics.

#### Scenario: Snapshot removed

- **WHEN** a compiled candlestick PPTX no longer contains the OfficeKit program
- **THEN** import returns the native group and no candlestick chart node

### Requirement: Agent guidance SHALL expose the primitive and its boundary

The generated PPJ manual and focused chart guide SHALL include the exact
channels, style fields, suitable use cases, editability and non-ChartPart
animation boundary.

#### Scenario: Fresh Agent searches for OHLC

- **WHEN** an Agent reads the chart reference
- **THEN** it can find a minimal candlestick example and the finite restrictions
