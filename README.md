
# RTM Analysis - Real-Time Managment

## Project Overview

This project provides a comprehensive Real-Time Monitoring (RTM) Analysis solution for workforce attendance, break adherence, and shift compliance. Built for Innovative Company, this Excel-based tool processes raw attendance data, work activity logs, and face recognition timestamps to generate actionable insights about agent performance and schedule adherence.

### Business Value

- Reduce unplanned absenteeism by 25 percent through early detection
- Improve break compliance by identifying adherence patterns
- Track logout events during active shifts to prevent time theft
- Generate compliance reports for client billing and audits

---

## Key Features

### Task 1 - Attendance Sheet Creation
- Automated population of login/logout times from Backend data
- Special handling for 7:00 AM and 12:00 AM shifts using Work Activity logs
- Calculation of working hours, missing time, and lateness
- Status classification (Active, Partial, OFF)

### Task 2 - Break Adherence
- Comparison of actual vs scheduled break durations
- Conditional formatting for visual compliance tracking:
  - Green: Within plus/minus 5 minutes of schedule
  - Orange: Short by more than 5 minutes
  - Red: Exceeds by more than 5 minutes

### Task 3 - OOA (Out of Adherence) for Breaks
- Detection of breaks taken during restricted periods:
  - First 30 minutes of shift
  - Last 30 minutes of shift
- Minute-precise OOA calculation using break timestamps

### Task 4 - Logout During Shift Detection
- Identification of premature logout events from Face ID system
- Duration calculation for each unauthorized logout
- Alert generation for compliance review

### Bonus Features
- Adherence Percentage Dashboard - Per-agent summary with pivot tables
- Planned vs Actual Chart - Visual comparison of break durations
- Productivity Metrics - Active hours vs total worked hours

---

## Technology Stack

| Component | Technology |
|-----------|------------|
| Primary Tool | Microsoft Excel (Office 365) |
| Data Processing | Excel Formulas (XLOOKUP, INDEX/MATCH, SUMIFS) |
| Array Operations | LET, FILTER, Dynamic Arrays |
| Visualization | PivotTables, Conditional Formatting, Charts |
| Data Sources | Excel Tables (.xlsx) |

---

## Sheet Descriptions

| Sheet Name | Purpose | Key Columns |
|------------|---------|-------------|
| Attendance | Master output sheet | Employee ID, Date, Scheduled Shift, Login/Logout, Break, Working Hours, Missing, Lateness, Exceeding Break |
| Backend | Raw attendance summary | clockIn, clockOut, Total Worked, Break, Idle, Active |
| Work Activity | Granular activity logs | Start/End Time, Activity, Worked for, Took a break |
| SCH | Scheduled shifts (wide format) | Employee Code, Date columns with shift times |
| Face ID | Biometric timestamps | Time, Personnel ID, Event Description, Device Name |
| Break Times | Extracted break periods | Break Start/End, Duration (min), OOA calculation |
| Adherence Summary | Bonus dashboard | Agent Name, Total Scheduled/Actual Break, Adherence Percent |

---

## Key Formulas Reference

### 1. Scheduled Shift Lookup (SCH Sheet - Wide Format)

`=IFERROR(INDEX('SCH'!$D$2:$ZZ$1000, MATCH(A2, 'SCH'!$A$2:$A$1000, 0), MATCH(D2, 'SCH'!$D$1:$ZZ$1, 0)), "OFF")`

### 2. Break Duration from Backend

`=IF(OR(E2="OFF", E2="Annual", E2="Sick"), "0:00:00", XLOOKUP(1, (Backend!B:B=A2)*(Backend!A:A=D2), Backend!H:H, "0:00:00"))`

### 3. Lateness Calculation (minutes)

`=MAX(0, (G2 - E2) * 1440)`

### 4. Break Adherence (minutes)

`= (I2 * 1440) - 30`

### 5. OOA (Out of Adherence) for Breaks

`=SUMIFS('Break Times'!E:E, 'Break Times'!A:A, A2, 'Break Times'!B:B, D2, 'Break Times'!C:C, "<"&(E2+TIME(0,30,0))) + SUMIFS('Break Times'!E:E, 'Break Times'!A:A, A2, 'Break Times'!B:B, D2, 'Break Times'!D:D, ">"&(E2+TIME(9,0,0)-TIME(0,30,0)))`

### 6. Logout During Shift Detection

`=IF(COUNTIFS('Face ID'!G:G, A2, 'Face ID'!A:A, "<"&(E2+TIME(9,0,0)), 'Face ID'!A:A, ">"&G2, 'Face ID'!E:E, "*out*")>0, "Yes", "No")`

### 7. Conditional Formatting - Break Adherence

| Color | Condition | Formula |
|-------|-----------|---------|
| Green | Within schedule | =AND(N2>=-5, N2<=5, N2<>0) |
| Orange | Short by more than 5 min | =N2<-5 |
| Red | Exceeds by more than 5 min | =AND(N2>5, N2<>0) |

---

## Dashboard Outputs

### Adherence Summary Table

| Agent Name | Scheduled Break (min) | Actual Break (min) | Adherence Percent |
|------------|----------------------|-------------------|-------------------|
| Abubasha | 30 | 38 | 78.9% |
| Ali | 30 | 42 | 71.4% |
| Hassan | 30 | 25 | 120.0% |
| Layla | 30 | 35 | 85.7% |

## Setup Instructions

### Prerequisites
- Microsoft Excel 2019 or later (Office 365 recommended)
- Enable Dynamic Arrays and LET functions
- Arizona Time zone setting (no DST)

### Installation Steps

1. Clone the repository
   `git clone https://github.com/yourusername/RTM-Analysis.git`

2. Open the main workbook
   `start [YourName]_RTM_Test.xlsx`

3. Enable macros if using VBA helper functions
   - File -> Options -> Trust Center -> Enable all macros

4. Refresh data connections
   - Data -> Refresh All

5. Verify formulas
   - Check that all XLOOKUP references point to correct sheets
   - Verify date formats match (MM/DD/YYYY)

---

## Performance Metrics

| Metric | Target | Current |
|--------|--------|---------|
| Break Adherence | plus/minus 5 min | 78% of agents |
| On-time Login | less than 5 min late | 92% compliance |
| Logout During Shift | 0 events | Detected and reported |
| OOA Breaks | 0 min | Calculated per agent |

---

## Troubleshooting

### Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| VALUE! | Text in time field (e.g., Annual, Casual) | Use IF(ISTEXT(...), OFF, ...) |
| N/A | No matching Employee ID + Date | Check that IDs match across sheets |
| NAME? | Function not recognized | Upgrade Excel or use alternative formula |
| SPILL! | Dynamic array blocked | Clear cells below formula |
| 0:00:00 for all breaks | Break column not found | Verify Backend column letter |

### Data Quality Checks

- Employee IDs consistent across all sheets
- Date formats unified (use DATEVALUE if needed)
- No merged cells in data ranges
- All time columns formatted as Time, not Text

## File Structure

```
RTM-Analysis/
│
├── README.md # Project documentation
├── [YourName]_RTM_Test.xlsx # Main analysis file
│
├── data/
│ ├── raw/
│ │ ├── Backend.xlsx # Raw attendance data
│ │ ├── Work_Activity.xlsx # Agent activity logs
│ │ ├── SCH.xlsx # Scheduled shifts
│ │ ├── Face_ID.xlsx # Face recognition logs
│ │ └── Agents_Breaks_Shifts.xlsx # Scheduled break data
│ │
│ └── processed/
│ ├── Break_Times.xlsx # Extracted break timestamps
│ └── Adherence_Summary.xlsx # Aggregated metrics
│
├── sheets/
│ ├── Attendance # Main output sheet
│ ├── Break_Times # Break extraction helper
│ ├── Adherence_Summary # Bonus summary table
│ └── ReadMe # Formula documentation
│
├── formulas/
│ ├── SCH_Lookup.txt # INDEX+MATCH formula
│ ├── OOA_Calculation.txt # Out of adherence logic
│ ├── Break_Adherence.txt # Comparison formulas
│ └── Conditional_Formatting.txt # Highlighting rules
│
├── charts/
│ ├── Planned_vs_Actual_Breaks.png # Comparison chart
│ └── Adherence_by_Agent.png # Performance dashboard
│
└── docs/
├── methodology.md # Analysis approach
├── assumptions.md # Data assumptions
└── troubleshooting.md # Common issues and fixes
```
## Author

**Mahmoud Abdallah**

### Any Questions
**Mahmoud_Abdallah20@outlook.com**

<img width="200" height="200" alt="logo" src="https://github.com/user-attachments/assets/c7b6037f-603f-407d-b42b-87824f7f2f75" />

