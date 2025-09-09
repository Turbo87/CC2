# ClaimCheck - Glider Flight Record Verification Software

ClaimCheck is an Excel-based application developed by Judy Ruprecht for verifying and processing glider (sailplane) flight record claims. The software automates the complex calculations and validations required for official glider flight records according to soaring competition and record-keeping standards.

## Overview

The software consists of 5 main Excel files with embedded VBA macros that work together to:
- Process and validate flight data from GPS trackers
- Calculate distances, speeds, and other performance metrics
- Verify compliance with official soaring rules and regulations
- Generate printable claim verification documents
- Manage waypoints and turn points for various flight tasks

## File Structure

### Excel Files
- **A.xlsm** - Main application interface and claim processing (214 KB)
- **Ab.xlsm** - Flight analysis and record calculations (182 KB)
- **C.xlsm** - Task verification and validation engine (666 KB)
- **D.xlsm** - Waypoint and turn point management (116 KB)
- **F.xlsm** - Flight data processing and analysis (425 KB)

### Extracted VBA Macros
The VBA source code has been extracted from each Excel file into corresponding directories:
- `A/` - Interface and workflow management modules
- `Ab/` - Record calculation and analysis functions
- `C/` - Task verification and rule validation
- `D/` - Waypoint management and calibration
- `F/` - Flight data import and processing

### User Interface Documentation
- **[ui_layout_reference.md](ui_layout_reference.md)** - Comprehensive documentation of all Excel worksheet layouts, user interface elements, and their corresponding cell references for cross-referencing with VBA macros

## Key Features

### Flight Record Types Supported
- **Distance Records**: Straight distance, out-and-return, triangle courses
- **Speed Records**: Various task types with official timing
- **Altitude Records**: Gain of height calculations
- **Duration Records**: Longest flight validation

### Data Processing Capabilities
- GPS flight log import and validation
- Pressure altitude corrections and calibrations
- Electronic vs. written pre-flight declarations
- Turn point sector verification
- Start/finish line crossing validation
- Ground track analysis and optimization

### Compliance Features
- SC3 soaring competition rules enforcement
- Official distance calculations using great circle methods
- Finish line crossing requirements for specific task types
- Minimum distance thresholds (e.g., 300km for triangles)
- ENL (Engine Noise Level) monitoring for motor gliders

## Technical Details

### Password Protection
The Excel files are password-protected with the password "spike" for worksheet protection.

### Key Functions
- **Claim verification**: Automated checking against official rules
- **Distance calculations**: Great circle and ground track computations
- **Time validation**: Start/finish timing verification
- **Data import**: Processing of various GPS file formats
- **Report generation**: Printable claim verification documents

### Integration
The modules work together through inter-workbook linking and shared data validation, with the main interface (A.xlsm) coordinating the overall workflow.

## Development Notes

- Original author: Judy L. Ruprecht (JLR)
- Development period: Multiple revisions from 2013-2018
- Platform: Microsoft Excel with VBA macros
- Architecture: Distributed across multiple workbooks for modularity

## Usage Context

This software is used by glider pilots, competition officials, and record verification committees to process and validate official soaring flight records according to international soaring regulations and competition rules.
