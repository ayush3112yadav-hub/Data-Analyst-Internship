# Week 2: Data Transformation and Pivot Table Analysis

Part of the **Data Analyst Internship** project, continuing Project FOODFLOW from Week 1. This week takes a public FAO dataset through the full workflow in Excel: import, cleaning, transformation and PivotTable analysis.

## Files in this folder
| File | What it is |
|------|------------|
| [Data_Transformation_and_Pivot_Table_Analysis.pdf](./Data_Transformation_and_Pivot_Table_Analysis.pdf) | The report (PDF, opens right in GitHub) |
| [Data_Transformation_and_Pivot_Table_Analysis.docx](./Data_Transformation_and_Pivot_Table_Analysis.docx) | The report (Word version) |
| [Food_supply_FAO_2020.xlsx](./Food_supply_FAO_2020.xlsx) | The Excel workbook: raw data, cleaned table and four PivotTables |
| [Entity.xlsx](./Entity.xlsx) | Entity classification list (Country, Aggregate or Historical) used to filter the data |

## What this covers
- **Dataset:** FAO food supply per person by country and year, 1961 to 2017 (11,028 rows, 18 columns).
- **Cleaning log:** nine issues found and resolved, each with the reasoning behind the decision.
- **Transformations and formulas:** a Type lookup, a blank-value flag and a meat-total consistency check, all explained.
- **Four PivotTables:**
  1. Average meat supply per person by year
  2. The meat mix by year, with a stacked column chart
  3. Top 10 countries by meat supply in 2017
  4. Year-on-year growth of average meat supply
- **Insights and limitations**, linked to the Week 1 research questions and hypotheses.

## Headline findings
- Average meat supply per person rose **77%** between 1961 and 2017, from 27.6 kg to 48.8 kg.
- **Poultry** grew from about 10% to 41% of meat supply, overtaking beef.
- Hong Kong's supply per person (137.1 kg) is about **36 times** India's (3.8 kg).
- 8,977 of the 11,028 rows (174 current countries) were analysed after removing regions, groups and historical states.

## Method in brief
1. Kept the original data untouched on a Raw_Data sheet and worked on a copy.
2. Converted the copy to an Excel Table and shortened the headers.
3. Classified every entity and filtered all pivots to **Type = Country**, so no country is counted twice.
4. Flagged blank pigmeat values instead of filling them.
5. Validated the data: the meat total equals the sum of its five meat types, with a largest difference of 0.16 kg.
6. Built the PivotTables and wrote an explanation and insight for each.

## Tools
Microsoft Excel (Tables, structured formulas, PivotTables, PivotCharts).

## Data source
FAO Food Balance Sheets (FAOSTAT), obtained through the Our World in Data repository on GitHub ("Food supply (FAO, 2020)"). Please credit FAO and Our World in Data if you reuse the data, and check their licence terms.

## Limitations
The data measures food **supply per person**, not factory output, so it is a demand signal for processors, not a measure of plant efficiency. Averages are unweighted, and the country list changes over time (see the report for details).

[Back to the main repository](../README.md)
