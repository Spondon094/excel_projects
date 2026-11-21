# Vehicle Crash Report Automation (Excel VBA)

**Full workbook:** [Google Drive](https://drive.google.com/drive/folders/1ratQWiDGjueSbPeWEpckXeO8IJEr2Sfx?usp=sharing)

Excel macros that build four monthly crash reports from one data sheet, each on its own sheet. This is a guided course project in Excel automation. The course scenario: a client needs the same report every month, so the reports should come from a macro instead of being built by hand.

**Files**

- `Vehicle_Crash_Report_Project.xlsx`: the workbook with the data and the four finished reports. It is saved as `.xlsx`, so it contains no macros.
- `crash_report_macros.bas`: the VBA code, as plain text.

## Data

7,879 vehicle crash records for New York City, all from January 2022 (1 to 31 January), with 29 columns: date, time, borough, street names, number of people injured and killed by type, contributing factors and vehicle types for up to five vehicles. Provided with the course. 2,701 records (34%) have no borough and 2,096 (27%) have no on-street name.

## The four reports

| Report | Sheet | Built with |
|---|---|---|
| Crashes by hour of the day | `Crash_times` | Pivot table grouped by hour, line chart |
| Unique street names where crashes happened | `Unique Addresses` | Remove duplicates, sort A to Z |
| Total injured and killed | `Total Injured and Dead` | Column totals |
| Crashes per contributing factor, highest first | `Count of Contributing_Factors` | Pivot table sorted descending |

## Findings

### Crashes by hour

- Crashes peak in the afternoon: 2 PM to 6 PM holds 2,345 crashes (29.8% of all), with 3 PM the highest hour (487) followed by 5 PM (484). There is a smaller morning peak at 8 AM (470).
- The quietest hour is 3 AM (134 crashes).
- The 12 AM bucket (332) is inflated. 118 crashes are logged at exactly 00:00, which looks like a default time. Without them the 12 AM hour has 214 crashes, in line with 1 AM (170) and 2 AM (178). In general the times look rounded: 18% of records fall on a full hour (minute :00).

### Injured and killed

| | Injured | Killed |
|---|---|---|
| All persons | 3,182 | 20 |
| Pedestrians | 673 | 10 |
| Cyclists | 174 | 1 |
| Motorists | 2,249 | 8 |

The three groups add up to 3,096 injured and 19 killed, so the "all persons" totals include 86 injured and 1 killed who are not in any of the three groups.

### Contributing factors (first vehicle)

| Factor | Crashes | Share |
|---|---|---|
| Unspecified | 1,982 | 25.2% |
| Driver Inattention/Distraction | 1,746 | 22.2% |
| Failure to Yield Right-of-Way | 559 | 7.1% |
| Following Too Closely | 461 | 5.9% |
| Pavement Slippery | 348 | 4.4% |

A quarter of crashes have no factor recorded ("Unspecified"). Among the recorded ones, driver inattention is far ahead of the rest.

### Unique street names

1,666 different on-street names appear in the data (compared without regard to case).

## How to run the macros

1. Open `Vehicle_Crash_Report_Project.xlsx` in Excel and save a copy as a macro-enabled workbook (`.xlsm`).
2. Press Alt+F11, then File > Import File, and choose `crash_report_macros.bas`.
3. Press Alt+F8 and run `Run_All_Reports`. To rebuild one report, run `Report_Crash_Times`, `Report_Unique_Addresses`, `Report_Total_Injured_Dead` or `Report_Contributing_Factors`.

Running a macro deletes the existing report sheet of the same name and rebuilds it. For a new month, replace the rows on the `Vehicle Crash Data` sheet with the new data, keeping the same column headers, and run the macro again.

## How the macros were written

The starting point was recorded macros, which contain fixed sheet names, fixed row counts (such as `7880`) and screen selections. They were rewritten so that:

- columns are found by their header text instead of their letter,
- the last row is taken from the data, so a longer or shorter month needs no edits,
- each report creates its own sheet, instead of assuming the sheet already exists,
- the pivot tables and the chart are built from variables rather than selections.

## Limitations

- The contributing-factor report counts only the first vehicle's factor. Factors for vehicles 2 to 5 are not included.
- The street report lists on-street names only. Cross streets and off-street addresses are not included, and the 2,096 records without an on-street name do not appear in it.
- The data covers one month, so the hourly pattern says nothing about seasons or weekdays.
- The analysis counts crashes and does not adjust for traffic volume.

## Skills used

Excel VBA (recording and refactoring macros, creating pivot tables and charts in code, removing duplicates, sorting), pivot tables, report automation.

## Repository

```
README.md
Vehicle_Crash_Report_Project.xlsx   workbook: Requirements, Vehicle Crash Data and the four report sheets
crash_report_macros.bas             VBA module that builds the four reports
```

## Author

Spondon Sarker, [github.com/Spondon094](https://github.com/Spondon094)
