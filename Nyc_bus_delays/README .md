# NYC School Bus Breakdowns and Delays (Excel)

**Full workbook:** [Google Drive](https://drive.google.com/drive/folders/1ratQWiDGjueSbPeWEpckXeO8IJEr2Sfx?usp=sharing)

Excel analysis of NYC school bus breakdown and delay records from September 2019 to December 2023, built as a guided project in a data analysis course. The course scenario: find patterns behind bus breakdowns and delays. It answers three questions with pivot tables and charts.

**Workbook:** `Bus_Breakdown_and_Delays_NYC_Project_lean.xlsx` (about 17 MB). The full workbook is about 70 MB, over what GitHub accepts for a normal upload, so this copy leaves out two sheets: the raw data and the `Cleaned` sheet (about 280,000 rows each). Everything else is unchanged: the scenario, the `Common Delay Reasons` sheet with its source rows, and the `Delay Times` and `Day of Week` sheets with their pivot tables and charts. The pivot tables on `Delay Times` and `Day of Week` show their saved results but cannot be refreshed, because their source sheet is not in this copy.

## Data

Incident records supplied with the course: 282,191 rows and 21 columns (run type, route, reason, time of incident, borough, bus company, reported delay range and others), school years 2019-20 to 2023-24 (up to 1 December 2023).

- 263,963 records are "Running Late" (93.5%) and 18,228 are "Breakdown" (6.5%).
- Volume differs a lot by year (20,438 records in 2020-21, 94,651 in 2022-23), so yearly counts are not comparable and the weekday results below are pooled across all years.

## Cleaning

In the full workbook the raw sheet is kept as it was, and the `Cleaned` sheet differs from it in four ways (neither sheet is in the copy in this repository):

- Bus company names were standardised in a second `Bus_Company_Name` column (77 distinct names became 54; 79,655 rows changed). The original column stays next to it, so the sheet has two columns with the same header.
- The delay range text (for example `46-60 Min`) was split into two numeric columns, `Short_Delayed_Time_Estimate` (46) and `Long_Delayed_Time_Estimate` (60).
- A `Day_Of_Week` column was added from the incident time.
- The `1900-01-01` placeholder in `Last_Updated_On` was blanked (252,833 rows).

## Questions and findings

### 1. What are the most common reasons for delays and breakdowns?

| Reason | Breakdowns | Share | Running late | Share |
|---|---|---|---|---|
| Mechanical Problem | 10,399 | 57.0% | 13,695 | 5.2% |
| Won't Start | 2,970 | 16.3% | 1,932 | 0.7% |
| Flat Tire | 2,499 | 13.7% | 2,401 | 0.9% |
| Other | 1,674 | 9.2% | 49,135 | 18.6% |
| Heavy Traffic | 279 | 1.5% | 176,196 | 66.8% |
| All other reasons | 407 | 2.2% | 20,604 | 7.8% |

- Mechanical Problem, Won't Start and Flat Tire together make up 87.1% of breakdowns. The data lists these as separate reasons, so treating them as one maintenance-related group is an interpretation.
- Heavy Traffic is the reason for two thirds of delays.
- The workbook recommends regular servicing for breakdowns and shifting pickups away from high-traffic times for delays. These are suggestions, and the analysis does not test them.

### 2. How do delay times vary by bus company and borough?

Delay times come from the reported delay range, as the lower and upper bound of each range. Breakdown records have no delay range, so only the 263,963 "Running Late" records are included. Across all of them the averages are 29.3 minutes (lower bound) and 45.1 minutes (upper bound).

**Ten companies with the highest average lower bound** (the workbook's top-10 filter):

| Company | Records | Avg lower bound (min) | Avg upper bound (min) |
|---|---|---|---|
| Selby Transportation | 908 | 59.9 | 88.1 |
| Little Linda Bus Co., Inc. | 95 | 51.7 | 75.2 |
| Pride Transportation (SCH AGE) | 19,221 | 49.0 | 70.4 |
| Little Richie Bus Service | 4,815 | 46.7 | 66.9 |
| Logan Bus Company Inc. | 3,585 | 46.3 | 65.8 |
| Fortuna Bus Company | 1 | 46.0 | 60.0 |
| IC Bus Inc. | 3 | 46.0 | 60.0 |
| MJT Bus Company, Inc | 18 | 45.2 | 60.0 |
| Lorinda Enterprises, Ltd. | 3,483 | 44.8 | 63.2 |
| Mar-Can Transport Co. Inc (B2192) | 2,397 | 44.5 | 65.6 |

Four of these ten companies have 95 records or fewer (three have 18 or fewer), so their averages say little. The well-supported results are Selby Transportation (908 records, the highest averages) and Pride Transportation (19,221 records, above the overall average).

**By borough:**

| Borough | Records | Avg lower bound (min) | Avg upper bound (min) |
|---|---|---|---|
| Manhattan | 61,611 | 37.0 | 53.1 |
| Rockland County | 448 | 33.3 | 51.1 |
| Queens | 44,002 | 32.9 | 49.7 |
| Staten Island | 17,738 | 32.5 | 46.9 |
| Westchester | 5,081 | 30.0 | 47.8 |
| New Jersey | 1,086 | 28.7 | 44.7 |
| Brooklyn | 62,663 | 27.3 | 42.9 |
| Nassau County | 13,238 | 26.2 | 41.4 |
| All Boroughs | 1,036 | 20.6 | 35.5 |
| Bronx | 56,777 | 20.0 | 35.3 |
| Connecticut | 196 | 11.7 | 26.2 |

Manhattan has the longest delays among the larger groups. The workbook explains this by population and traffic density. That explanation is a hypothesis, and the table only partly supports it: the Bronx has the shortest delays of the five NYC boroughs, and Staten Island is above Brooklyn.

### 3. Do delays or breakdowns depend on the day of the week?

| Day | Running late | Breakdowns | Total | Breakdown share | Avg incidents per date with records |
|---|---|---|---|---|---|
| Monday | 54,667 | 4,146 | 58,813 | 7.1% | 406 |
| Tuesday | 52,332 | 3,619 | 55,951 | 6.5% | 335 |
| Wednesday | 53,589 | 3,573 | 57,162 | 6.3% | 346 |
| Thursday | 53,862 | 3,572 | 57,434 | 6.2% | 357 |
| Friday | 49,511 | 3,318 | 52,829 | 6.3% | 330 |

- Monday has the most incidents even though it has the fewest dates with records (145, against 160 to 167 for the other days), and the highest share of breakdowns.
- Friday has the lowest total. Per date with records it is close to Tuesday (330 against 335), so the Friday dip is small.
- After Monday, the breakdown share is flat at 6.2% to 6.5%. The pattern is a higher Monday, not a steady decline through the week.
- The workbook explains the lower Friday figure by fewer people driving to work, and suggests weekend maintenance for the Monday breakdowns. Both are untested ideas; the analysis only counts incidents.

## Limitations

- The analysis is descriptive. There are no significance tests, so "difference" here means a difference in counts or averages.
- Delay times are averages of reported ranges, not measured minutes. One delay record has no reported range and is counted as 0 to 0.
- The weekday tables leave out 2 of the 282,191 records: one Sunday record and one record with an empty `Day_Of_Week`.
- 87 delay records have no borough and are missing from the borough table.
- Reasons are as reported in the source data.

## Skills used

Excel: pivot tables, top-10 filters, pivot charts, data cleaning, working with a 280,000-row dataset.

## Repository

```
README.md
Bus_Breakdown_and_Delays_NYC_Project_lean.xlsx   workbook: Scenario, Common Delay Reasons, Delay Times, Day of Week
```

## Author

Spondon Sarker, [github.com/Spondon094](https://github.com/Spondon094)
