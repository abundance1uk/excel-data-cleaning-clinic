# Healthcare Appointment Records: Data Cleaning, Quality Control & Dashboard

An end-to-end Excel project that takes a messy clinic appointment dataset, cleans and validates it, flags incomplete records, and presents the results in a Quality Report and an interactive dashboard.

![Dashboard](images/dashboard.png)

---

## The Problem

"Greenfield Community Clinic" (a fictional clinic) has an appointment records file full of errors: duplicate rows, inconsistent spellings, dates in four different formats, impossible ages and missing values. Records like these cannot be trusted for reporting, so the clinic needs them cleaned, validated and monitored.

**Dataset:** 2,150 synthetic appointment records (no real patient data), 8 columns: Patient ID, Name, Age, Gender, Appointment Date, Department, Appointment Status, Staff ID.

## What I Did

1. **Measured the problems first.** I counted every issue before changing anything, using `COUNTBLANK`, `COUNTIF` and `SUMPRODUCT`, so the results can be proved.
2. **Kept the raw data untouched.** All cleaning was done on a copy (`Cleaned Data`).
3. **Removed duplicates** with Remove Duplicates, then checked that Patient IDs were unique.
4. **Standardised text** with `TRIM` and `PROPER`, and Find & Replace with exact matching for Gender typos.
5. **Standardised Departments** with an `XLOOKUP` mapping table (for example "Paediatrics" → "Pediatrics", "ANC" → "Antenatal Care"), including a fallback value to catch anything unmapped.
6. **Converted mixed date formats** (text and real dates) into one `yyyy-mm-dd` format. Because Excel was set to US date order, day/month/year dates were rebuilt explicitly with `DATE`, `LEFT`, `MID` and `RIGHT` instead of letting Excel guess, which would have silently swapped days and months.
7. **Cleared impossible ages** (-3, 150, 200, 999) rather than guessing replacements.
8. **Flagged missing data, not invented it.** An alert column using `IF` and `TEXTJOIN` labels each incomplete record (for example "Missing Age, Missing Staff ID"), and conditional formatting highlights those rows in red.
9. **Added data validation** (dropdown lists for Gender, Department and Status, and an age range of 0-120) so the same errors can't be entered again. Excel's "Circle Invalid Data" check found zero problems in the cleaned data.
10. **Looked up staff names** with `XLOOKUP` against a Staff List, and built summaries with `COUNTIFS`.
11. **Built pivot tables, pivot charts and slicers**, then assembled them into a dashboard.

## Results

| Issue | Before | After |
|---|---|---|
| Duplicate records | 150 | 0 |
| Inconsistent Gender entries | 119 | 0 |
| Inconsistent Department entries | 72 | 0 |
| Dates stored as text (mixed formats) | 1,860 | 0 |
| Impossible ages | 23 | 0 (cleared and flagged) |
| Total records | 2,150 | 2,000 |
| Records with missing data | not tracked | 257 flagged (12.9%) |

**Missing fields across the 257 flagged records:** Age 115, Gender 42, Appointment Date 37, Status 37, Staff ID 39 (some records are missing more than one).

## Key Findings

- **Overall no-show rate: 16.9%** (337 of 2,000 appointments).
- **Completed appointments: 60.8%** (1,216 of 2,000).
- No-show rates vary by department, from **10.1% in Dental** (the lowest) to **21.3% in Laboratory** and 21.1% in HIV/TB Clinic (the highest).
- Age was the most commonly missing field (115 records).

## Dashboard

The dashboard has four KPI cards (Total Appointments, Completed %, No-Show Rate, Records Flagged), three charts (appointments by status, no-show rate by department, appointments by month) and slicers for Department, Month and Gender. One click updates every card and chart.

![Dashboard filtered to Dental](images/dashboard-dental.png)

## Files

| File | Description |
|---|---|
| `Greenfield_Clinic_Data_Management_Project.xlsx` | The full workbook |
| `Greenfield_Clinic_Dashboard.pdf` | Dashboard export |
| `Quality_Report.pdf` | One-page data quality report |
| `images/` | Screenshots |

**Workbook sheets:** `Dashboard`, `Monthly Summary`, `Cleaned Data`, `Calc` (pivot tables behind the dashboard), `Quality Report`, `README`, `Raw Data` (untouched original), `Staff List`, `Lookup Lists`.

## Skills Demonstrated

- **Data cleaning:** duplicate removal, text standardisation, date conversion, handling missing and invalid values
- **Excel functions:** `XLOOKUP`, `IF`, `COUNTIF`, `COUNTIFS`, `TEXTJOIN`, `TRIM`, `PROPER`, `DATE`, `SUMPRODUCT`
- **Data quality control:** data validation, conditional formatting, alert flags, before/after measurement
- **Reporting:** pivot tables, pivot charts, slicers, KPI dashboard design, PDF export
- **Documentation:** a clear audit trail from raw data to final report

## Decisions Worth Noting

- **No guessing.** Missing ages, genders and dates were flagged, not filled in, because invented values are worse than missing ones in health records.
- **Raw data was never edited,** so every step can be verified.
- **Validation as prevention,** not only cleanup.

## What I'd Do Next

- Add a staff attendance table and a monthly activity report.
- Rebuild the cleaning steps in Power Query so they repeat automatically on new data.

---

*Dataset is synthetic and contains no real patient information.*
