# Sales Data Analysis (Excel)

## Overview
A practice project analysing a small set of sales records using Excel.
The data is assumed (made up for practice), not from a real business.

## Skills demonstrated
- Data cleaning: TRIM and PROPER to fix extra spaces and inconsistent
  capital letters in employee names
- Formulas: IF, SUMIFS and calculated columns (Total Revenue = Units Sold x Price)
- Excel Tables: structured references such as Table1[Total Revenue]
- PivotTables: summarising sales by region (and by employee)

## Example formulas
Clean names:
=PROPER(TRIM(B2))

Total revenue per sale:
=E2*F2

Target check (units sold of 15 or more):
=IF(E2>=$K$1, "Yes", "No")

Revenue for Apple in the North region:
=SUMIFS(Table1[Total Revenue], Table1[Product], "Apple", Table1[Region], "North")

## Key findings
- South had the highest units sold (35), followed by North (17) and East (8)
- Total revenue across all sales was R105.50

## Note
All figures are fictional and used only to practise Excel skills.
