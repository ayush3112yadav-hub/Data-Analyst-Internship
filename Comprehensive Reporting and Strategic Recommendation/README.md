# Week 5: Comprehensive Reporting and Strategic Recommendation

The final task of the **Data Analyst Internship**. This report combines the planning, analysis, automation and dashboard work from Weeks 1 to 4 into one document, and turns it into **strategic recommendations** for NutriLine Foods, a virtual food processing company (Project FOODFLOW).

## Files in this folder
| File | What it is |
|------|------------|
| [Comprehensive_Reporting_and_Strategic_Recommendation.pdf](./Comprehensive_Reporting_and_Strategic_Recommendation.pdf) | The final report (PDF, opens right in GitHub) |
| [Comprehensive_Reporting_and_Strategic_Recommendation.docx](./Comprehensive_Reporting_and_Strategic_Recommendation.docx) | The final report (Word version) |

## The question
Where should a virtual food processor focus its products and capacity, and how should it work with data so that decisions are fast, accurate and repeatable?

## What the report covers
- **Executive summary:** the main findings and six recommendations on one page
- **Introduction and background:** the scenario, industry context and a table of the four-week journey
- **Methodology:** data sources, cleaning decisions, the advanced Excel techniques used in each week, and how results were validated
- **Analysis and findings:** six charts and supporting tables built from the FAO data
- **The interactive dashboard:** its components and how a manager uses it
- **Interpretation:** what the results mean for production quality, efficiency and cost
- **Strategic recommendations:** six recommendations, each with the finding, the action, KPIs, how Excel shaped it, and a caution
- **Conclusion and future directions:** summary, future analyses and ways to improve the data framework
- **Appendices:** key formula reference, project files and a plain-language glossary

## Headline findings
- Average meat supply per person rose **77%** between 1961 and 2017 (27.6 kg to 48.7 kg), and **poultry** grew from 9.8% to 41.3% of the meat supply, overtaking beef.
- Markets are very unequal: Hong Kong has about **36 times** India's meat supply per person (137.1 kg against 3.8 kg).
- The share of countries in severe food deficit fell from **45.2% to 4.7%** (1961 to 2017), but 8 countries are still flagged and 33 more are on a watch list.
- Algeria's milk supply rose about **3 times** and Brazil's poultry supply about **28 times**.

## The six recommendations
| # | Recommendation |
|---|----------------|
| R1 | Prioritise poultry processing capacity |
| R2 | Build a shelf-stable dairy offer for markets leaving food scarcity |
| R3 | Sort markets into three tiers and set targets by market, not by the global average |
| R4 | Run the annual early-warning list and plan affordable, fortified products |
| R5 | Make data-quality checks and automation standard practice |
| R6 | Collect plant-level data to measure real efficiency |

## How advanced Excel shaped the strategy
PivotTables and PivotCharts found the patterns, formulas (INDEX/MATCH, nested IF, SUMPRODUCT, SUMIFS, AVERAGEIFS, COUNTIFS, LARGE) and a VBA macro made them repeatable, the dashboard (named ranges, dropdowns, conditional formatting) made them easy to explore, and independent checks made them trustworthy.

## The whole project
| Week | Task | Folder |
|------|------|--------|
| 1 | Data Analysis Strategy and Planning | [Week 1](../Week-1-Data-Analysis-Strategy-and-Planning) |
| 2 | Data Transformation and Pivot Table Analysis | [Week 2](../Data-Transformation-and-Pivot-Table-Analysis) |
| 3 | Advanced Excel Functions and Automation | [Week 3](../Advanced-Excel-Functions-and-Automation) |
| 4 | Dashboard Creation and Data Visualisation | [Week 4](../Dashboard-Creation-and-Data-Visualisation) |
| 5 | Comprehensive Reporting and Strategic Recommendation | This folder |

## Data source
FAO Food Balance Sheets (FAOSTAT), obtained through the Our World in Data repository on GitHub ("Food supply (FAO, 2020)"). Please credit FAO and Our World in Data if you reuse the data.

## Important note on what is real and what is illustrative
The analysis of FAO data is real. The efficiency index, hypothesis tests H1 to H3 and all dollar figures from the Week 1 plan are **illustrative examples** that show how the method works, and the report labels them as such. The data measures food **supply** per person, not food eaten or factory output.

[Back to the main repository](../README.md)
