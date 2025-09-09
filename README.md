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

## Entry Points and Key Functions

### Main Entry Point
- **`OpenAb()`** in [A/Module148.bas](A/Module148.bas) - Primary entry point for flight analysis workflow
  - Handles IGC file selection and opening
  - Initiates flight data processing chain
  - Sets up user interface for verification tasks

### Core Processing Functions
- **`NewENLA()`** in [Ab/Module1.bas](Ab/Module1.bas) - Engine Noise Level processing and coordinate analysis
- **`NewBRecords()`** in [Ab/Module1.bas](Ab/Module1.bas) - B-record parsing and coordinate extraction
- **`NEWHilo()`** in [Ab/Module1.bas](Ab/Module1.bas) - High/low point analysis and distance calculations
- **`RefineLDG()`** in [Ab/Module1.bas](Ab/Module1.bas) - Landing point refinement for high-density flights

### Typical Workflow
1. User clicks glider interface element to trigger `OpenAb()`
2. System prompts for IGC file selection
3. IGC file is parsed and transferred to analysis workbook (Ab.xlsm)
4. `NewENLA()` processes flight data and extracts coordinates
5. `CALx()` performs flight calculations and verification
6. User interface displays results for verification and review

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