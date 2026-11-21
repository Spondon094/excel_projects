# LAPD Crime Dashboard (Excel, Power Query and Data Model)

**Full workbook:** [Google Drive](https://drive.google.com/drive/folders/1ratQWiDGjueSbPeWEpckXeO8IJEr2Sfx?usp=sharing)

An Excel dashboard of crime reports to the Los Angeles Police Department, built as a guided course project in Excel automation. The course scenario: the data arrives as one CSV file per year, so the workbook should load every file in a folder by itself and update when a new file is added, instead of being copied and pasted by hand.

**The workbook (about 216 MB) is not in this repository.** GitHub rejects files over 100 MB. This repository holds the query code, the description and screenshots of the dashboard.

## Data

Crime reports published by the LAPD, one CSV file per year (2020 to 2024), 28 columns each: report number (`DR_NO`), report and occurrence date, area, crime code and description, victim age, sex and descent, premises, weapon, status and location. The files are not included here; see the LAPD open data portal for the source.

| Year | Reports |
|---|---|
| 2020 | 192,708 |
| 2021 | 208,284 |
| 2022 | 235,065 |
| 2023 | 234,644 |
| 2024 (partial) | 103,776 |
| **Total** | **974,477** |

2024 stops at 2 September 2024, so it is not comparable with the full years. No report number (`DR_NO`) appears twice across the five files.

## How the automation works

1. **Power Query, "From Folder".** One query reads every CSV in a folder, applies the same transform to each file (comma-delimited, header row promoted, 28 columns), and appends them into one table. A new year is added by dropping the CSV into the folder and refreshing. The code is in `lapd_power_query.pq`.
2. **Cleaning in the query.** Columns get explicit types (report number as whole number, coordinates as numbers). `Date Rptd` and `DATE OCC` are converted from datetime to date. Repeated spaces in `LOCATION` are collapsed to one.
3. **Data Model.** The combined table is loaded into the Excel Data Model (Power Pivot) rather than only onto a sheet, so the pivot tables run on all rows. A sheet holds at most 1,048,575 data rows, and the 974,477 reports leave about 74,000 spare.
4. **Pivot tables and charts.** Four pivot tables on a staging sheet feed the charts and the headline number on the `Dashboard` sheet.

## Dashboard

| Visual | Source | Shows |
|---|---|---|
| Headline number | Pivot on the Data Model | Total reports, 974,477 |
| Reports by year and month | Pivot grouped by report date | Monthly counts per year |
| Victims by sex | Pivot on `Vict Sex` | Reports per category |
| Weapon used | Pivot on `Weapon Desc` | Weapon types, filtered to 2,000 to 550,000 reports each |

## Findings

All figures are counts of reports, grouped by report date.

### Over time

- Reports rose from 192,708 in 2020 to 235,065 in 2022 (+22.0%), then were flat in 2023 (234,644, -0.2%).
- The highest month is May 2022 (20,597). The lowest full month is November 2020 (15,342).
- Counts for the partial 2024 year fall from 19,902 in January to about 8,300 a month in June to August. The data does not say why; it may reflect how the file was extracted, so this drop should not be read as fewer crimes.

### Victims by sex

| Sex | Reports | Share |
|---|---|---|
| M | 394,539 | 40.5% |
| F | 350,930 | 36.0% |
| Not recorded | 136,003 | 14.0% |
| X | 92,893 | 9.5% |
| Other codes (H, "-") | 112 | 0.0% |

Sex is "not recorded" for 14% of reports, and the source documentation for what `X` means should be checked before it is interpreted. The chart shows reports by victim sex, not a rate per population.

### Weapons

The weapon chart keeps weapon types with between 2,000 and 550,000 reports. This hides reports with no weapon recorded (the largest group, above the upper limit) and weapon types below 2,000. It covers 301,601 reports.

- Strong-arm (hands, fists, feet or bodily force) accounts for 174,296 of them (57.8%).
- Unknown or other weapon: 35,799 (11.9%). Verbal threat: 23,786 (7.9%).
- Hand gun: 20,076 (6.7%). Adding semi-automatic pistol (7,240) and unknown firearm (6,564) gives 33,880 firearm-related reports (11.2%). That grouping is mine; the data lists them separately.

## How to rebuild it

1. Put the CSV files in their own folder, with nothing else in it (no workbook, no other file types). Excel's folder query reads every file, including a saved workbook, and fails with `The column 'DR_NO' of the table wasn't found`.
2. In Excel: Data > Get Data > From File > From Folder, point to the folder, then Transform Data. Or open the Advanced Editor and paste `lapd_power_query.pq`, with the folder path replaced.
3. Close & Load To... > "Only Create Connection" and tick "Add this data to the Data Model".
4. Insert pivot tables from the Data Model and build the charts.

To update with a new CSV, drop it in the folder and use Data > Refresh All. If a table linked into Power Pivot does not update, use Power Pivot > Linked Table > Update All.

## Limitations

- 2024 is a partial year (to 2 September), and its monthly counts fall off from April, so year totals must not be compared with 2024.
- Reports are grouped by report date (`Date Rptd`), not by the date the crime occurred.
- If two files overlap, the same report is loaded twice. The query does not remove duplicates; with these five files there are none.
- The weapon chart is filtered (see above) and does not show reports without a recorded weapon.
- One chart title in the workbook has a typo ("LYPD" instead of "LAPD").
- The analysis is descriptive and counts reports. It does not adjust for population or reporting practices.

## Skills used

Excel Power Query (folder import, combining files, type conversion, text cleaning), Data Model and Power Pivot, pivot tables and pivot charts, dashboard layout, handling a dataset near the Excel row limit.

## Repository

```
README.md
lapd_power_query.pq     Power Query code (folder path is a placeholder)
images/                dashboard screenshots
```

## Author

Spondon Sarker, [github.com/Spondon094](https://github.com/Spondon094)
