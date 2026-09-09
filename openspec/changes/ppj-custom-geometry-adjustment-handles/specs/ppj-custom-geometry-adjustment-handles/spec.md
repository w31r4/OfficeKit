## Purpose

Preserve XY and polar custom-shape adjustment handles as typed PPJ state with editable ranges and positions.

## ADDED Requirements

### Requirement: Typed handles preserve native controls

Custom shapes SHALL accept up to 1024 ordered adjustmentHandles of kind xy or polar. XY SHALL expose xAdjustment/minX/maxX and yAdjustment/minY/maxY; polar SHALL expose radialAdjustment/minRadius/maxRadius and angleAdjustment/minAngle/maxAngle. Both SHALL require position {x,y}. At least one declared adjustment MUST be controlled. Coordinates/radii SHALL use local points and angles degrees, or native reference strings. Paired range presence, zero, reference identity and order SHALL survive export/fresh projection. Empty authored lists SHALL project as omission. Invalid ownership, ranges, references and positions MUST reject; masks/clips MUST reject handles.

#### Scenario: XY and polar records
- **WHEN** a custom shape declares mixed literal/reference XY and polar handles
- **THEN** exported and freshly projected records SHALL preserve their values and presence

### Requirement: Source handle edits retain control identity

Issued setGeometry authority SHALL permit adjustmentHandles position/range changes while retaining source order, kind and controlled adjustment names. Identity and list-length changes MUST reject. Source no-op SHALL preserve bytes; edits SHALL retain unchanged geometry/text/frame and non-target ZIP members. Paired bounds SHALL support addition, replacement and removal within that identity.

#### Scenario: Range lifecycle
- **WHEN** separate original-source requests add, change or remove paired bounds and change positions
- **THEN** projection SHALL recover the requested values while retaining the original controls

### Requirement: Preview reports handle limitations

Preview SHALL diagnose handle fields it cannot render rather than reporting complete host interaction support.

#### Scenario: Unrendered handle
- **WHEN** a custom shape declares an adjustment handle
- **THEN** unsupported handle state SHALL remain explicitly diagnosed
