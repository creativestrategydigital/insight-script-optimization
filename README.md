# Audit Route Generator

A Node.js script that generates optimized audit routes from SR (Sales Representative) visit data, distributing visits across auditors while respecting geographic proximity and business constraints.

## Features

- **D+2/D+3 Constraint**: Each visit can only be audited 2-3 days after the SR's visit date
- **Geographic Optimization**: Routes are optimized using nearest-neighbor algorithm to minimize travel distance
- **Even Distribution**: Audits are distributed fairly across all auditor dates and individual auditors
- **Buffer Management**: Creates a buffer sample for replacements (120 outlets)
- **SR Fair Distribution**: Ensures each SR has a proportional number of audits based on their universe size
- **Excel Export**: Generates formatted Excel files with routing information

## Prerequisites

- Node.js v14+
- npm packages: `xlsx`

```bash
npm install xlsx
```

## Usage

```bash
node generate_audit_routes.js "input_file.csv" > audit_plan.txt
```

### Input File Format

The input CSV must contain the following columns:

| Column | Description |
|--------|-------------|
| `DATE` | Visit date (DD/MM/YYYY format) |
| `SEM ID` | Unique identifier for the visit |
| `DB-ID` | Database ID |
| `SR` | Sales Representative name |
| `Territory` | Territory name |
| `Outlet Name` | Outlet name |
| `Region` | Region name |
| `New Channel` | Channel type |
| `Telephone` | Contact number |
| `Latitude` | Geographic latitude |
| `Longitude` | Geographic longitude |

### Output Files

The script generates two Excel files:

1. **Audit_Main_400.xlsx** - Main audit routes (400 audits)
   - `Auditor`: Auditor number (1-4)
   - `AuditDate`: Date of audit
   - `Sequence`: Visit order for the day
   - `OriginalVisitDate`: SR's original visit date
   - `SR`: Sales Representative
   - `Territory`, `Outlet`, `SEM_ID`, etc.

2. **Audit_Buffer_120.xlsx** - Buffer sample (120 outlets)
   - Same structure as main file
   - Used for replacements if needed

## Configuration

Edit these constants at the top of the script:

```javascript
// Auditor dates (when auditors are available)
const AUDITOR_DATES = [
  "18/05/2026", "19/05/2026", "20/05/2026",
  "21/05/2026", "22/05/2026", "23/05/2026",
  "25/05/2026", "26/05/2026", "27/05/2026",
  "28/05/2026", "29/05/2026", "30/05/2026"
];

const NUM_AUDITORS = 4;           // Number of auditors per day
const MIN_VISITS_PER_AUDITOR = 8; // Minimum visits per auditor per day
const MAX_VISITS_PER_AUDITOR = 10; // Maximum visits per auditor per day
const TARGET_AUDITS = 400;        // Total main sample size
const BUFFER_SIZE = 120;          // Buffer sample size
```

## Algorithm Overview

### 1. Data Loading & Validation
- Reads CSV file
- Parses coordinates and dates
- Filters out invalid entries (missing coordinates)

### 2. Eligibility Check (D+2/D+3)
- For each visit, calculates eligible audit dates (visit date + 2 days, + 3 days)
- Only keeps visits that fall within the auditor availability window

### 3. SR Grouping
- Groups visits by Sales Representative
- Each SR gets proportional audits based on their total visits
- Ensures fair distribution across SRs

### 4. Sample Selection
- **Main Sample**: 400 audits selected proportionally from SRs
- **Buffer Sample**: 120 additional audits for replacements

### 5. Daily Target Calculation
- Calculates fixed daily targets (400 ÷ 12 days ≈ 33-34 audits/day)
- Validates that each date has enough eligible visits
- Adjusts targets if eligibility is insufficient

### 6. Date Assignment
- Randomizes main sample
- Assigns each visit to one of its eligible dates
- Respects daily capacity limits

### 7. Route Optimization
For each audit date:
1. Sorts visits by latitude (north to south)
2. Distributes evenly across 4 auditors
3. Optimizes each auditor's route using nearest-neighbor algorithm
4. Assigns sequence numbers for visit order

### 8. Export
- Generates Excel files with formatted data
- Includes summary statistics

## Console Output

The script outputs detailed logs:

```
✓ Total visits loaded: 522
✓ Eligible visits: 402
✓ Main sample: 400
✓ Buffer sample: 120

======================
FIXED DAILY TARGETS
======================
18/05/2026 : 34 audits
19/05/2026 : 34 audits
...

======================
ELIGIBLE VS TARGETS
======================
✓ 20/05/2026: TARGET=34, ELIGIBLE=61
⚠ 18/05/2026: TARGET=34 but ELIGIBLE=19 → ADJUSTED TO 19

======================
FINAL RESULTS
======================
✓ Main audits exported: 400
✓ Buffer exported: 120
✓ Average audits/day: 33.50
✓ Average audits/auditor/day: 8.38
```

## Warnings

The script will display warnings if:
- An auditor has fewer than 8 visits on a given day
- An auditor has more than 10 visits on a given day
- A date doesn't have enough eligible visits to meet its target

## Example

```bash
node generate_audit_routes.js "Abidjan_Mai_16_28.csv" > audit_plan.txt
```

This will:
1. Load visits from `Abidjan_Mai_16_28.csv`
2. Generate optimized routes
3. Create `Audit_Main_400.xlsx` and `Audit_Buffer_120.xlsx`
4. Save the execution log to `audit_plan.txt`

## License

Internal use only.
