# Excel VBA Application UI Documentation

This document describes the user interfaces found in the Excel VBA files, detailing worksheet layouts, controls, and their cell references for cross-referencing with VBA macros.

## File Overview

The application consists of 5 main Excel workbooks:
- **A.xlsm** - Main application with claims processing
- **Ab.xlsm** - Data processing and parsing workbook  
- **C.xlsm** - Claims verification and task analysis
- **D.xlsm** - Way points and calibration management
- **F.xlsm** - Additional task calculations

---

## A.xlsm - Main Claims Application

### Sheet: ALL CLAIMS (Main Interface)
**Purpose**: Primary data entry and control interface for flight claims

#### Key UI Elements:

| Row | A | B | C | D | E | F | G |
|-----|---|---|---|---|---|---|---|
| 1 |   |   |   |   |   |   |   |
| 2 | 28 |   | v 2.9 |   | ✓ |   |   |
| 3 | 29 |   |   |   |   |   |   |
| 4 |   |   | BASIC FLIGHT DATA |   |   | ALL applications | [START/CLEAR] |
| 5 |   |   |   |   |   |   |   |
| 6 |   |   | Default units: ____ |   |   | ✓ |   |
| 7 |   |   |   |   |   |   |   |
| 8 |   |   | Aircraft Designation: ____ |   |   | ✓ |   |
| 9 |   |   |   |   |   |   |   |
| 10 |   |   | Region of Takeoff & Time Zone: |   | ✓ | ✓ |   |
| 11 |   |   |   |   |   |   |   |
| 12 |   |   |   | ✓ |   |   |   |
| 13 |   |   |   |   |   |   |   |
| 14 |   |   | [Dynamic Field 1] |   |   |   |   |
| 15 |   |   |   |   |   |   |   |
| 16 |   |   | [Dynamic Field 2] |   |   |   |   |
| 17 |   |   |   |   |   |   |   |
| 18 |   |   | Declaration Type: ____ |   |   | ✓ |   |

**Control Elements:**
- **G4**: `START / CLEAR` button - Main action control
- **F6**: Default units selector (cell reference for VBA)
- **F8**: Aircraft designation toggle (cell reference for VBA) 
- **F10,D12**: Region/timezone selectors
- **F18**: Declaration type selector

**Dynamic Content Areas:**
- **C6-C18**: Form labels that change based on selections
- **C14,C16**: Dynamic fields based on unit selection (F6=2)

#### Selection Areas (Bottom of sheet):

| Row | C | F | G |
|-----|---|---|---|
| 39 | Select One | Select One | Select One |
| 40 | Electronic | Metric Default | Pure Glider SOLO |
| 41 |   |   | Pure Glider 2 Pilots |
| 42 |   |   | Motorglider 2 Pilots |

### Sheet: E-DEC (Electronic Declaration)
**Purpose**: Electronic flight declaration interface

#### Layout:

| Row | B | C | D | E | F | G | H | I |
|-----|---|---|---|---|---|---|---|---|
| 4 | Electronic Declaration |   | ✓ |   |   |   |   |   |
| 6 | FR & Serial # | [Formula] |   |   | Flight Date | [Date] |   |   |
| 8 | Pilot(s) | [Formula] |   | Aircraft |   | [Formula] |   |   |
| 10 | E- Declaration Date & Time (UTC): | [Formula] |   |   | # Declared Turnpoints | [#] |   |   |
| 12 |   | Latitude | Degrees | Minutes | Longitude | Degrees | Minutes |   |
| 13 | Start Point | [Name] | [Lat] | [Min] | [Lon] | [Min] |   |   |
| 15 | Turnpoint A | [Name] | [Lat] | [Min] | [Lon] | [Min] |   |   |
| 17 | Turnpoint B | [Name] | [Lat] | [Min] | [Lon] | [Min] |   |   |
| 19 | Turnpoint C | [Name] | [Lat] | [Min] | [Lon] | [Min] |   |   |
| 24 | Click a tab below to continue |   |   |   |   |   |   |   |

**Key References:**
- **I10**: Number of turnpoints (max 3 validation at F4)
- **C13,C15,C17,C19**: Turnpoint names
- **E13,F13,H13,I13**: Start point coordinates
- **E15,F15,H15,I15**: Turnpoint A coordinates
- Similar pattern for turnpoints B and C

### Sheet: OTHER (Custom Declarations)
**Purpose**: Custom and written declaration interface

#### Key Controls:

| Row | B | D | F | G | H | I | J | K | L | M | N |
|-----|---|---|---|---|---|---|---|---|---|---|---|
| 4 | Evaluation Basis: | [1] |   |   |   |   |   |   |   |   |   |
| 15 | Coordinate Format: | [1] | [2] |   |   |   |   |   |   |   | SELECT ONE |
| 16 |   |   |   |   |   |   |   |   |   |   | YES |
| 17 | [Dynamic Help Text] |   |   |   |   |   |   |   |   |   | NO |
| 18 | UTC Declaration Date &Time: |   |   | # of Turn Points: [#] |   |   |   |   |   |   |
| 19 | Place Name | Lat,Deg | MM | [SS/.mmm] | N or S | Long,Deg | MM | [SS/.mmm] | E or W |   |   |
| 20 | Start Point | [____] | [__] | [____] | [_] | [____] | [__] | [____] | [_] |   |   |

**Interactive Elements:**
- **D4**: Evaluation basis selector
- **D15,F15**: Coordinate format selectors  
- **K15,N15**: Additional option selectors
- **N18**: Yes/No selector for saved waypoints

### Sheet: DATA ENTRY CHECK
**Purpose**: Flight data validation and verification

#### Main Interface:

| Row | C | D | E | F | G | I |
|-----|---|---|---|---|---|---|
| 4 | DATA ENTRY CHECK: |   |   |   |   |   |
| 6 | [File Type] | [FR Info] | Flight Date | [Date] | [Aircraft] | [Crew Info] |
| 8 | [Declaration Info] | [Decl Type] |   | [W/OO] | Flight Crew | [Crew Data] |
| 10 | [Data Point Info] | [Time] | [Status] | [Ref Time] | [Landing Info] | [Time] |
| 12 | Confirm or Correct Release Time here: | [Time] | hh:mm:ss or hh:mm | [enter] |   |

**Key Interaction Points:**
- **G12**: Release time correction field (editable)
- **I12**: Time format help text
- **C34**: Navigation instruction to ALL CLAIMS tab

### Sheet: Logo
**Purpose**: Application branding/status

---

## Ab.xlsm - Data Processing Workbook

### Sheet: PRS (Processing Sheet)
**Purpose**: Flight record parsing and processing

#### Key Elements:

| Row | A | B | C | D | E | F | G | H | I | J | K |
|-----|---|---|---|---|---|---|---|---|---|---|---|
| 1 |   |   | FR (mfr, #) | Pre-flt PA |   | End T/O Roll | ABS HI | FLARM |   |   |
| 2 |   |   | Date; UTC | TO elev(1) |   | EndSL |   |   |   |   |
| 3 |   |   |   | Calc RelAlt |   | RelUTC Calc |   | LXN Red Box Flarm |   |   |
| 4 |   | Pilot(s) | RelTime | Rel Lat | HI gain |   |   |   |   |   |
| 8 | User Pilot(s) |   |   |   |   |   |   |   |   |   |
| 9 |   | End in Flight |   |   |   |   |   |   |   |   |
| 14 | [Release Time Field] |   |   |   |   |   |   |   |   |   |

**Dynamic Processing Areas:**
- **C3-D12**: Calculation formulas referencing imported data
- **A4-B5**: Source code handling for pilot information  
- **A8**: Combined pilot names for display

---

## C.xlsm - Claims Verification System

### Sheet: SUMMARY (Flight Summary)
**Purpose**: Consolidated flight information display

#### Layout:

| Row | A | B | C | D | E | F | G | H | I | J | K | L |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Claim Check IGC File Summary |   |   |   |   |   |   |   |   |   |   |   |
| 4 | [FR Type] | [Info] | [#] | FR Model |   | UTC REL | REL ALT |   |   |   | VSTART Duration |   |
| 6 | Sailplane Type |   | [#] | Last Calibration |   | Rel Lat Deg | Rel lat Min |   |   |   |   |   |
| 8 | Sailplane ID |   | [#] | UTC LOW | LO ALT |   | Rel Long Deg |   | Straight Distance |   |   |   |
| 9 | Flight Date |   |   |   |   |   | [Coordinates continue...] |   |   |   |   |   |
| 13 | Declaration Date & Time |   |   |   |   |   |   |   |   |   |   |   |
| 22 | Start Name | [Coords spread across multiple columns] |   |   |   |   |   |   |   |   |   |
| 25 | Start Long Deg | Start Long Min | Goal Start |   |   |   |   |   |   |   |   |   |

### Sheet: Worksheet (Main Data Entry)
**Purpose**: Primary applicant data entry interface

#### Key Sections:

| Row | B | C | D | E | F | G | H | I | J | K | L | M | N | O |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | APPLICANT DATA ENTRY |   |   | CALCULATED VALUE | BADGE & RECORD CLAIM WORKSHEET |   |   |   |   |   |   |   |   |   |
| 2 |   |   | SCROLL DOWN TO VIEW DISTANCE CLAIM DATA |   |   |   |   |   |   |   |   |   |   |   |
| 3 | Pilot Name | [Name] |   |   |   |   |   |   |   |   | Flight Date | [Date] |   |   |
| 5 | Data Recorder: Make/Model |   |   |   | Serial # | [#] | Cal Due | [Date] |   |   |   |   |   |   |
| 7 | Aircraft | [Aircraft Info] |   |   |   |   |   |   |   |   |   |   |   |   |
| 11 | Release | Dur Start Fix | Start Line |   | Low Point | ABS High | High Gain | Finish Line |   |   |   |   |   |   |
| 12 | [Coords] | [Times and measurements across columns] |   |   |   |   |   |   |   |   |   |   |   |   |

**Control Elements:**
- **C11-O11**: Duration and line controls
- **N1**: Recorder status (formula-driven)
- **B32**: Geodesic distance calculations section

### Sheet: VERIFY TASK (Task Verification)
**Purpose**: Flight task verification interface

#### Interactive Elements:

| Row | A | B | C | D | E | F | G | H | I |
|-----|---|---|---|---|---|---|---|---|---|
| 2 | [X] | [Task Info] |   | Calculated |   | Altitude Basis | [Date] | Select One |   |
| 4 |   | DISTANCE & SPEED APPLICANTS: | Turn Points Achieved & Turn Point Order |   |   |   | Pressure Data |   |   |
| 8 |   | Select Altitude Basis above, then proceed to CALIBRATION |   |   |   |   |   | GPS Data |   |
| 9 |   |   |   | Calculated TP Use Order | Confirm or Correct below |   |   |   |   |
| 10 |   |   |   | [Turnpoint Data] |   |   |   |   |   |
| 12 | [TP1] | [Distance] | [Status] | [Use Order] |   |   |   |   |   |
| 14 | [TP2] | [Distance] | [Status] | [Use Order] |   |   |   |   |   |
| 16 | [TP3] | [Distance] | [Status] | [Use Order] |   |   |   |   |   |
| 29 |   |   |   |   |   |   |   |   | END / EXIT |

**Key Controls:**
- **I2**: Altitude basis selector (Pressure Data/GPS Data)
- **F12,F14,F16**: Turn point correction fields
- **H29**: Exit control

---

## D.xlsm - Waypoints & Calibration

### Sheet: Saved Way Points
**Purpose**: Waypoint management interface

#### Main Interface:

| Row | A | B | C | D | E | F | G | H | I | J | K | L |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|
| 2 |   | For Written Declarations: ADD A SAVED WAY POINT |   |   |   |   |   |   |   |   |   |   |
| 3 |   |   |   |   |   |   |   | [Status Message] |   |   |   |   |
| 4 |   | Coordinate Format | [Selector] |   |   |   |   |   |   |   |   |   |
| 7 |   | Unique Name | Latitude |   |   | Longitude |   |   |   |   |   |   |
| 8 |   | CASE SENSITIVE | Degrees | Minutes | .mmm | Seconds | N or S | Degrees | Minutes | .mmm | Seconds | E or W |
| 9 |   | [____] | [___] | [___] | [___] | [___] | [_] | [___] | [___] | [___] | [___] | [_] |
| 13 |   | Other editing must be done one Way Point at a time... |   |   |   |   |   |   |   |   |   |   |
| 16 | [1] | [Name] | [Coordinate data across columns] |   |   |   |   |   |   |   |   |   |
| 17 | [2] | [Name] | [Coordinate data across columns] |   |   |   |   |   |   |   |   |   |
| 42 |   | Click on the glider to continue |   |   |   |   |   |   |   |   |   |   |

**Interactive Elements:**
- **B4**: Coordinate format selector
- **B9-L9**: New waypoint entry fields
- **I3**: Dynamic status messages
- **N9**: Validation status ("OK" when complete)

### Sheet: Calibration (Calibration Data)
**Purpose**: Flight recorder calibration management

#### Interface Layout:

| Row | A | B | C | D | E | F | G | H | I | J | K | L | M |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2 |   | CALIBRATION DATA |   |   | Data Entry Required |   |   |   | Start / Clear |   |   |   |   |
| 6 |   |   |   |   |   |   |   |   |   | [Status Messages] |   |   |   |
| 8 |   | [FR Count Status] |   |   | [Entry Type] |   |   |   | Start / Clear |   |   |   |   |
| 10 |   | [FR Info] | [Cal Data] |   | FR Serial #: [____] |   |   |   |   |   |   |   |   |
| 12 |   |   | [Cal Date] |   | Last Calibration Date: [Date] | FR MANUFACTURER CODE |   |   |   |   |   |   |
| 16 |   | [FR Details when selected] |   |   |   |   |   |   |   |   |   |   |   |
| 48 |   | Click on the glider to save |   |   |   |   |   |   |   |   |   |   |   |

**Key Controls:**
- **I8,I28**: Start/Clear buttons for different sections
- **E10**: Entry type selector (2-5 different modes)
- **E12**: FR serial number entry
- **E14**: Last calibration date
- **L6-L8**: Status and action messages

---

## F.xlsm - Task Calculations

### Sheet: TASKS
**Purpose**: Task-specific calculations interface

### Sheet: YDWK3 (Geodesic Calculations)  
**Purpose**: Advanced distance and bearing calculations

#### Calculation Areas:

| Row | B | C | D | E | F | G | H | I | J | K | L | M | N | O | P | Q |
|-----|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Ellipsoid | WGS84 |   | OZ FOR O&R FINI |   |   |   |   |   |   |   |   |   |   |   |   |
| 2 | Azimuth 1-2 (a12) | [Calc] | [+45°] |   |   |   |   |   |   |   |   |   |   |   |   |   |
| 3 | Azimuth 2-1(a21) | [Calc] | [-45°] |   |   |   |   |   |   |   |   |   |   |   |   |   |
| 5 | Lat1, Lon1 RAD | [Coordinates and calculations continue...] |   |   |   |   |   |   |   |   |   |   |   |   |   |   |

---

## Cross-Reference with VBA Macros

### Key VBA-to-UI Connections:

#### UserForm Interactions:
1. **frmUNIT.frm**: Unit Selection Form (D/Module6.bas:RoundedRectangle12_Click)
   - Triggered from D.xlsm Calibration sheet controls
   - Purpose: Unit selection for calibration procedures
   
2. **frmCAL.frm**: Calibration Form (D/Module6.bas:RoundedRectangle1_Click)
   - Triggered from D.xlsm Calibration sheet controls  
   - Purpose: Calibration data entry and management

#### Workbook Integration (A/Module1.bas):
1. **OpenD()** Subroutine:
   - **Trigger Cells**: A.xlsm OTHER sheet D15, K15 (coordinate format selectors)
   - **Actions**: 
     - When D15=1: Simple cell selection
     - When D15>1 AND K15≠2: Clears coordinate entry area (C69:O93)
     - When K15=2: Opens D.xlsm for waypoint management
   - **UI Elements Controlled**: 
     - Shapes ("Oval 14", "Oval 16", "Oval 18", "Oval 20", "Oval 22", "Rectangle 1")
     - Coordinate entry area visibility

2. **SeeList()** Subroutine:
   - Opens D.xlsm Saved Way Points sheet
   - Calls D.xlsm!List macro
   - Sets focus to D4 (coordinate format selector)

#### Workbook Initialization (A/ThisWorkbook.cls):
1. **Workbook_Open()** Event:
   - **Internet Connectivity Check**: Updates A2 query table if connected
   - **Link Management**: Updates links to D.xlsm calibration data
   - **UI Configuration**: 
     - Disables formula bar and status bar
     - Sets full screen mode
     - Configures calculation settings
   - **Data Initialization**: 
     - Clears Parsed sheet A39:A64
     - Populates C40:L64 with calibration data from D.xlsm

#### Cell Reference Patterns:
```
A.xlsm ALL CLAIMS:
- F6: Unit selector → Triggers label changes in C6-C18
- F8: Aircraft designation → Controls C8 label visibility  
- F10, D12: Region/timezone → Controls C10 label content
- F18: Declaration type → Major UI mode changes
- G4: START/CLEAR button → Primary action handler

A.xlsm OTHER:
- D4: Evaluation basis → Controls multiple UI sections
- D15, K15: Coordinate format → Triggers OpenD() workflow
- N15, N18: Yes/No selectors → Controls waypoint functionality

C.xlsm VERIFY TASK:
- I2: Altitude basis selector → Controls calculation mode
- G2: Task verification status → Referenced across multiple sheets

D.xlsm Calibration:  
- E8: Entry mode selector → Controls form display
- E10: FR selector → Populates calibration data
```

#### Cross-Workbook References:
The application uses extensive cross-workbook linking:
- **[4]PRS!**: References to Ab.xlsm PRS sheet (parsed flight data)
- **[1]OTHER!**: References to A.xlsm OTHER sheet (declaration data)  
- **[2]VERIFY TASK!**: References to C.xlsm verification data
- **'[D.xlsm]Calibration'**: Direct links to calibration data

#### Dynamic UI Control Mechanisms:
1. **Conditional Visibility**: Shapes and controls hidden/shown based on selections
2. **Formula-Driven Labels**: Cell contents change based on other cell values
3. **Protection Management**: Sheets protected/unprotected programmatically
4. **Window State Control**: Workbooks minimized/maximized as needed
5. **Cross-Sheet Navigation**: Automatic sheet switching based on workflow

### Common VBA Interaction Patterns:
- **Worksheet_Change events**: Monitor cells F6, F8, F10, F18 in ALL CLAIMS
- **Click event handlers**: Process buttons and shape clicks
- **Cross-workbook communication**: Opening/closing linked workbooks
- **Data synchronization**: Copying values between sheets and workbooks  
- **UI state management**: Controlling visibility and protection of elements

---

## Navigation Flow

### Primary User Workflow:
1. **A.xlsm ALL CLAIMS**: Main data entry and declaration type selection
2. **A.xlsm E-DEC/OTHER**: Specific declaration type interfaces  
3. **C.xlsm Worksheet**: Detailed applicant data entry
4. **C.xlsm VERIFY TASK**: Task verification and turnpoint management
5. **A.xlsm DATA ENTRY CHECK**: Final validation
6. **C.xlsm PRINT THIS!**: Output generation

### Secondary Workflows:
- **D.xlsm**: Waypoint and calibration management (accessed from OTHER sheet)
- **Ab.xlsm**: Background data processing (typically hidden from user)
- **F.xlsm**: Advanced calculations (typically hidden from user)

This UI structure supports a complex flight claims processing system with multiple declaration types, coordinate formats, and validation steps.