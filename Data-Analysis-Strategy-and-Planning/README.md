# Week 1: Data Analysis Strategy and Planning

**Deliverable:** [Project FOODFLOW: Data Analysis Plan](./Project_FOODFLOW_Data_Analysis_Plan.docx)

## What this covers

### 1. Project introduction and context of virtual food processing
Virtual food processing means studying and optimising the food-processing chain through data and simulation instead of on a physical shop floor: a "digital twin lite" of a plant, tested in an Excel sandbox. The project models a fictional mid-sized plant, **NutriLine Foods**, benchmarked against real public industry data. Data-driven decisions matter because of thin margins, volatile inputs, perishability and safety risks, and rising pressure on energy use and sustainability.

### 2. Objectives, six research questions and four hypotheses
**Objectives**
- **O1:** Consolidate at least five public datasets into one cleaned Excel data model (10+ years where available).
- **O2:** Quantify production trends, seasonality and volatility across five food categories.
- **O3:** Build a Production Efficiency Index (PEI).
- **O4:** Test four hypotheses using correlation, regression and lag analysis.
- **O5:** Deliver an interactive dashboard and five or more recommendations, each with a what-if value.

**Research questions**
1. How have production volumes moved over 10 to 15 years, and which categories are the most volatile?
2. Which segments create the most value added per worker and per unit of energy?
3. Do raw-material price changes squeeze processing margins, and after what delay?
4. Does higher capital intensity lead to higher productivity?
5. Where is NutriLine Foods' biggest efficiency gap, and what is closing it worth?
6. What does the next three years look like under base, optimistic and stressed scenarios?

**Hypotheses**
- **H1:** Capital-intensive segments produce more output per worker.
- **H2:** Raw-material price spikes reduce margins with a lag of one to two periods.
- **H3:** Segments with higher energy use per tonne score lower on overall efficiency.
- **H4:** Staple categories (cereals, oils) are more stable than perishable ones (fruit and vegetables, dairy).

### 3. Scope and limitations
- **In scope:** public data on production, prices and manufacturing performance; India, US and EU comparison; five food categories; a fictional benchmark plant; all work done in Excel.
- **Out of scope:** proprietary plant data, real-time IoT data, machine-learning models, lab and compliance audits, financial modelling of real companies.
- **Limitations:** data is aggregated rather than plant-level, definitions differ across sources, official data is published with a lag, Excel has row limits, and correlation does not prove causation.

### 4. Five public data sources with the rationale for each
| Source | Why it was chosen |
|---|---|
| **FAOSTAT** (UN FAO) | The global standard for production volumes, yields and producer prices, consistent across countries. |
| **US Census Annual Survey of Manufacturers** | The richest public view of labour, capital, energy and materials in food manufacturing. |
| **India Annual Survey of Industries / data.gov.in** | Adds an Indian lens and a developing versus developed economy comparison. |
| **Eurostat** (SBS and PRODCOM) | A third regional benchmark with harmonised definitions across EU countries. |
| **Open Food Facts** | A product-mix layer for profiling segments by product type. |

Selection logic: official sources first, triangulate across regions, free and Excel-friendly downloads, and confirm exact tables and licences before use.

### 5. Advanced Excel methodology, including a custom Production Efficiency Index
The analysis runs like a processing line with five stations: **Intake, Inspect, Process, Package, Dispatch**. Data quality is protected by five HACCP-inspired gates covering completeness, consistency, plausibility, traceability and repeatability.

The **Production Efficiency Index (PEI)** combines labour productivity, energy intensity and yield into one 0 to 1 score. Each component is normalised and weighted (for example 40% / 30% / 30%), with the weights held in named cells so the ranking's sensitivity can be tested.

**Excel toolkit:** Power Query, Excel Tables, XLOOKUP, SUMIFS, dynamic arrays (FILTER, SORT, UNIQUE, LET), PivotTables with slicers, conditional formatting, data validation, CORREL and LINEST, FORECAST.ETS, Data Tables, Scenario Manager, Goal Seek and Solver.

### 6. 33-hour timeline and resource plan
| Phase | Hours |
|---|---|
| P1. Frame and plan | 2 |
| P2. Intake: find and collect data | 5 |
| P3. Inspect and clean | 5 |
| P4. Process: model and indicators | 4 |
| P5. Analyse and test | 7 |
| P6. Package: dashboard | 4 |
| P7. Dispatch: insights and report | 4 |
| P8. Quality assurance and buffer | 2 |
| **Total** | **33** |

Planned as 10 working sessions of about three hours over two to three weeks. **Resources:** Excel 365 with the Data Analysis ToolPak and Solver enabled, five free public datasets, one analyst plus a peer reviewer, and no software or data cost.

### 7. Risk assessment with mitigation strategies
Each risk is scored on likelihood and impact (1 to 3), and anything scoring 6 or more gets active mitigation from day one.

| Risk | Mitigation |
|---|---|
| Inconsistent definitions across sources (priority 9) | Build a category mapping table early and compare growth and ratios, not just absolute values. |
| Missing or delayed data | Set a 5% completeness threshold and exclude failing series rather than guess. |
| Spreadsheet formula errors | Use Tables, named ranges and check cells, plus a peer audit. |
| Currency and inflation distortions | Convert to one currency and use inflation-adjusted values for trends. |
| Over-reading correlations | Present findings as evidence, test across sources, and show r-squared. |
| Large files slow Excel | Filter and aggregate in Power Query before loading. |
| Scope creep and time overrun | Hold to five questions and five sources, and use the 2-hour buffer. |
| Data licence or access changes | Keep dated copies of downloads and record licence terms in the README. |
