# US Public Debt Analysis (Excel)

Excel analysis of US public debt from April 1993 to February 2023, built as a guided project in a data analysis course. It answers three questions with pivot tables, charts and a forecast.

**File:** `US_Debt_Tracker_Project.xlsx`

## Data

Daily US public debt figures supplied with the course, about 7,500 records, with three measures:

- Debt Held by the Public
- Intragovernmental Holdings
- Total Public Debt Outstanding (the sum of the two)

The raw data was reshaped into a table (`Raw Data` to `Cleaned Data`) before analysis.

## Questions and findings

### 1. Yearly percentage increase in debt

Year-end to year-end change in Total Public Debt Outstanding:

| Period | Yearly increase |
|---|---|
| 2008 / 2009 / 2010 | 15.9% / 15.1% / 13.9% |
| 2014 to 2019 | 2.6% to 7.2% per year |
| 2020 | 19.6% (the largest yearly increase between 2005 and 2022) |
| 2021 / 2022 | 6.7% / 6.1% |

The 2020 jump coincides with the pandemic period. The workbook does not test its cause.
The 2023 figure (0.11%) covers only 1 January to 15 February 2023, so it is not comparable with full years.

### 2. Monthly pattern in Total Public Debt Outstanding

Average Total Public Debt Outstanding by calendar month (pivot table, all years in the data):

| Month | Avg (USD trillion) | Month | Avg (USD trillion) |
|---|---|---|---|
| Jan | 13.09 | Jul | 12.59 |
| Feb | 12.89 | Aug | 12.66 |
| Mar | 12.67 | Sep | 12.74 |
| Apr | 12.41 | Oct | 12.81 |
| May | 12.43 | Nov | 12.97 |
| Jun | 12.49 | Dec | 13.04 |

Overall average: 12.73 trillion. The highest averages are in January, December and November, and the lowest in April and May.

Note: these are average debt levels, not increases. The data runs from April 1993 to February 2023, so January and February averages include 2023 and leave out 1993, while April to December do the opposite. This affects the comparison between months.

### 3. Projected growth of publicly held debt

Forecast of the yearly maximum of Debt Held by the Public, using Excel's `FORECAST.ETS` (exponential smoothing) on 1997 to 2022:

| Year | Projected (USD trillion) |
|---|---|
| 2023 | 25.9 |
| 2024 | 28.4 |
| 2025 | 30.2 |
| 2026 | 31.0 |
| 2027 | 33.0 |

Method limitation: from 2024 onward, each forecast formula includes the earlier forecast values in its input range, so later years build on earlier predictions. Treat the 2024 to 2027 figures as indicative.

## Skills used

Excel: pivot tables, date grouping, `FORECAST.ETS`, absolute references, charts.

## Repository

```
US_Debt_Tracker_Project.xlsx   workbook: Scenario, Raw Data, Cleaned Data, Yearly Percentage Change,
                               Monthly Historicals, Projected Growth, Final Outputs
README.md
```

## Author

Spondon Sarker, [github.com/Spondon094](https://github.com/Spondon094)
