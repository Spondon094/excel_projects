# Excel Projects

Four Excel projects from guided courses in data analysis and Excel automation: two analyses built with pivot tables and charts, and two automated workflows built with VBA macros and Power Query. Each project has its own folder with a README that gives the data, method, findings and limitations.

**Full workbooks:** [Google Drive](PASTE-LINK-HERE). Some workbooks are too large for GitHub (the LAPD one is about 216 MB), so they are on Drive.

## Projects

### Data analysis

| Project | Folder | What it covers | Tools |
|---|---|---|---|
| US Public Debt Analysis | [`us-debt-analysis`](us-debt-analysis/) | Daily US public debt, April 1993 to February 2023: yearly increase, monthly pattern and a forecast of publicly held debt | Pivot tables, date grouping, `FORECAST.ETS`, charts |
| NYC School Bus Breakdowns and Delays | [`nyc-bus-delays`](nyc-bus-delays/) | 282,191 incident records, 2019 to 2023: common reasons, delay times by company and borough, day-of-week pattern | Pivot tables, top-10 filters, pivot charts, data cleaning |

### Automation

| Project | Folder | What it covers | Tools |
|---|---|---|---|
| Vehicle Crash Report Automation | [`vehicle-crash-report-automation`](vehicle-crash-report-automation/) | Macros that build four monthly crash reports from one data sheet (7,879 NYC crash records) | VBA, pivot tables and charts in code, remove duplicates, sorting |
| LAPD Crime Dashboard | [`lapd-crime-dashboard`](lapd-crime-dashboard/) | A dashboard that loads every yearly CSV in a folder (974,477 reports, 2020 to 2024) and updates when a file is added | Power Query (from folder), Data Model, pivot tables, charts |

## Summaries

**US Public Debt Analysis.** Total public debt grew 19.6% in 2020, the largest yearly increase between 2005 and 2022. The forecast for 2023 to 2027 uses exponential smoothing, and from 2024 on it builds on earlier forecast values, so those figures are indicative only.

**NYC School Bus Breakdowns and Delays.** Mechanical Problem, Won't Start and Flat Tire make up 87.1% of breakdowns. Heavy Traffic is the reason for 66.8% of "Running Late" records. Monday has the highest share of breakdowns (7.1%). The analysis is descriptive, and the explanations in the workbook are not tested.

**Vehicle Crash Report Automation.** Recorded macros with fixed sheet names and row counts were rewritten to find columns by header, take the last row from the data and create their own report sheets, so the same code works for a new month. In the January 2022 data, crashes peak between 2 PM and 6 PM (29.8% of all).

**LAPD Crime Dashboard.** One Power Query reads all CSV files in a folder, applies the same cleaning to each and loads the result into the Data Model, so adding a new year is a drop-in file plus Refresh All. Reports rose from 192,708 in 2020 to 235,065 in 2022 and were flat in 2023. 2024 is a partial year, to 2 September.

## Repository structure

```
excel_projects/
  README.md
  us-debt-analysis/
  nyc-bus-delays/
  vehicle-crash-report-automation/
  lapd-crime-dashboard/
```

## About these projects

- They are guided course projects. The scenarios and datasets come from the courses; the READMEs state what was done and where findings are limited.
- Findings are descriptive counts and averages. There are no significance tests, and explanations of causes are marked as untested.
- GitHub rejects files over 100 MB, so the workbooks for the larger projects are on Drive. The LAPD folder holds the Power Query code and screenshots instead.
