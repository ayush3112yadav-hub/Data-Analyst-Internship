# Week 4: Dashboard Creation and Data Visualisation

Part of the **Data Analyst Internship** project, continuing Project FOODFLOW. This week turns the cleaned FAO data from Weeks 2 and 3 into an **interactive Excel dashboard** that a manager can explore without touching a formula.

## Files in this folder
| File | What it is |
|------|------------|
| [Dashboard_Creation_and_Data_Visualisation.pdf](./Dashboard_Creation_and_Data_Visualisation.pdf) | The report (PDF, opens right in GitHub) |
| [Dashboard_Creation_and_Data_Visualisation.docx](./Dashboard_Creation_and_Data_Visualisation.docx) | The report (Word version) |

## The dashboard: Global Food & Nutrition Analytics
A one-screen dashboard for NutriLine Foods, a virtual food processor, built on FAO food supply data for 174 countries from 1961 to 2017.

**Interactivity:** three dropdowns (Country, Year, Metric) drive every number, chart and sentence on the screen.

**Visual components**
- **Five KPI cards:** metric value, country average, difference from average, rank among countries, and nutritional status.
- **Line chart:** the selected country against the average of all countries, 1961 to 2017.
- **Bar chart:** the top 10 countries for the chosen metric and year.
- **Doughnut chart (a pie variant):** the meat mix (beef, poultry, pork, mutton, other) for the selected country and year.
- **Key Insights box:** three plain-language sentences that rewrite themselves for each selection.

**Excel techniques used**
- Excel Table (tblFood) with helper columns that switch the metric with INDEX/MATCH
- Named ranges for the dropdown lists
- SUMIFS, AVERAGEIFS and COUNTIFS for KPIs, and LARGE with INDEX/MATCH for the top 10
- Dynamic headings built from formulas, and "No data" handling for missing years
- Conditional formatting: red/green difference card, traffic-light status card, arrow icons, top-10 highlight and data bars

## What the report covers
1. Introduction and objective
2. Data selection
3. Dashboard design plan (audience, questions, layout)
4. Key performance indicators and why they were chosen
5. Step-by-step build with screenshots
6. Visual components and why each chart type was chosen
7. Interactivity and dynamic ranges
8. Conditional formatting
9. Insights and practical implications for food processing
10. Testing and validation
11. Limitations and improvements
- Appendices: formula reference and a rebuild checklist

## Headline insights
- **Poultry is growing fastest.** Brazil's poultry supply rose from 1.73 kg to 48.03 kg per person (1961 to 2017), about 28 times higher, while the country average rose 7.4 times.
- **Transition markets grow dairy demand.** Algeria's milk supply rose from 38.94 kg to 121.71 kg per person, and its nutritional status moved from Severe Deficit to Adequate.
- **Markets are diverging.** Afghanistan's meat supply fell from 52% of the country average in 1961 to 18% in 2017.
- **Markets sit in tiers.** In 2017, Belgium had 3,768 kcal per person per day against 1,758 kcal in the Central African Republic.
- **Local mix matters.** India has the lowest meat supply of 172 countries but an above-average milk supply.

## Data source
FAO Food Balance Sheets (FAOSTAT), obtained through the Our World in Data repository on GitHub ("Food supply (FAO, 2020)"). Please credit FAO and Our World in Data if you reuse the data.

## Limitations
The data measures food **supply** per person, not food eaten. "Country average" gives every country equal weight, and "global rank" means rank among the 146 to 172 countries with data in that year. The status card is based on calories only. See the report for details.

[Back to the main repository](../README.md)
