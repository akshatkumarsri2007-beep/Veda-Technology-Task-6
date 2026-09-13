# Veda-Technology
# Task 6 — Excel Formulas & Functions Fundamentals

**Internship:** Veda Technology — Data Analytics Track
**Level:** 1 · Day 6 of 45
**Submitted by:** Akshat Srivastava

# Overview

This task is about building fluency with the everyday Excel formulas used in
real analyst work — VLOOKUP, INDEX/MATCH, IF, SUMIFS, COUNTIFS, and text
functions — practiced on a small transactional dataset.

# Objective

Turn a set of core Excel functions into a working reference: a workbook where
every formula runs on live data (not hardcoded results) and recalculates
correctly, plus short notes on when to reach for each one.

# Dataset

A custom transactional dataset created for this task — **15 orders, 10
fields**: Order ID, Order Date, Customer Name, Region, Category, Product,
Quantity, Unit Price, Amount, Payment Status. The `Amount` column is a live
formula (`Quantity × Unit Price`), and the `Customer Name` column was left
intentionally messy (extra spaces, inconsistent casing) to give the text
functions a real cleaning problem to solve.

A separate `Price_List` reference table (Product → Category → Unit Price →
Reorder Level) supports the lookup examples.

# What's in the workbook

| Sheet | What it demonstrates |
|---|---|
| `Transactions` | Base dataset used by every other sheet |
| `Price_List` | Lookup/reference table |
| `VLOOKUP_demo` | Category, price, and reorder level pulled by product name |
| `INDEX_MATCH_demo` | Same lookups solved with INDEX + MATCH — the compatible stand-in for XLOOKUP |
| `IF_demo` | Single-condition IF (High Value flag) and nested IF (Small/Medium/Large tiers) |
| `SUMIFS_COUNTIFS_demo` | Totals and counts by region, and by category + payment status combined |
| `TEXT_functions_demo` | TRIM, PROPER, UPPER, and LEFT/MID/FIND used to clean and split names |
| `Notes` | One-page cheat sheet — when to use each function and common pitfalls |

# Why INDEX/MATCH instead of XLOOKUP

XLOOKUP isn't supported consistently across every spreadsheet application
(older Excel, Google Sheets, LibreOffice can all handle it differently). This
workbook uses `INDEX` + `MATCH` instead, which does the same two-way lookup
job and stays compatible everywhere the file might be opened.

# Key outputs

- South region produced the highest total sales, followed by West.
- 4 of the 15 orders (~27%) qualified as "High Value" (₹2,000+).
- Electronics + Paid was the largest revenue combination under a
  two-condition SUMIFS filter.
- Every messy raw name was normalised correctly by `TRIM` + `PROPER`, and
  first/last name splitting held up even with leading/trailing spaces once
  cleaning ran first.

# Challenges & learnings

- **XLOOKUP compatibility** — swapped for INDEX/MATCH so the file opens
  correctly everywhere.
- **Cleaning order matters** — `FIND`-based name splitting broke until
  `TRIM` ran first; parsing before cleaning silently produces wrong results
  instead of an error.
- **Nested IF boundaries** — tier cutoffs (500 and 2,000) needed care so no
  order landed in the wrong bucket.
- SUMIFS/COUNTIFS scale cleanly to multiple AND conditions — the default
  choice over SUMPRODUCT for everyday filter-and-total work.

# Files in this repo

| File | Description |
|---|---|
| `Excel_Formulas_Fundamentals.xlsx` | The workbook — all functions demonstrated on live, recalculating data |
| `Task6_Report.pdf` | Full write-up: approach, components, outputs, challenges, learnings |
| `README.md` | This file |

# Tools
Microsoft Excel (workbook built to also open correctly in Google Sheets).
