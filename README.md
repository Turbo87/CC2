# ClaimCheck - Glider Flight Record Verification

ClaimCheck is an Excel-based application by Judy Ruprecht for verifying glider flight record claims according to official soaring standards.

## Architecture

The software consists of 5 Excel workbooks with VBA macros:

| File | Purpose | VBA Modules |
|------|---------|-------------|
| [A.xlsm](A.xlsm) | Main interface and workflow coordination | [A/](A/) |
| [Ab.xlsm](Ab.xlsm) | Flight analysis and record calculations | [Ab/](Ab/) |
| [C.xlsm](C.xlsm) | Task verification and rule validation | [C/](C/) |
| [D.xlsm](D.xlsm) | Waypoint and turn point management | [D/](D/) |
| [F.xlsm](F.xlsm) | Flight data import and processing | [F/](F/) |

## Documentation

- **[UI Layout Reference](ui_layout_reference.md)** - Complete worksheet layouts and cell references for cross-referencing with VBA code

## Supported Record Types

- **Distance**: Straight, out-and-return, triangle courses
- **Speed**: Various tasks with official timing  
- **Altitude**: Gain of height calculations
- **Duration**: Flight time validation

## Key Capabilities

- GPS log import and validation
- SC3 soaring rules compliance
- Great circle distance calculations
- Pressure altitude corrections
- Turn point sector verification
- ENL monitoring for motor gliders
- Automated claim verification reports

## Technical Notes

- **Password**: "spike" for worksheet protection
- **Author**: Judy L. Ruprecht (2013-2018)
- **Platform**: Microsoft Excel with VBA
- **Users**: Glider pilots, officials, record committees