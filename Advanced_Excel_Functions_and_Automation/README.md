# Project FOODFLOW: Advanced Excel Functions & VBA Automation

An automated annual nutritional-supply early-warning system built on historical UN FAO food supply data (1961–2017). Developed for **NutriLine Foods** to screen global food balance sheets, evaluate severe caloric shortfalls, and generate dynamic action reports with one-click VBA automation.

---

## 📌 Project Overview

* **Author:** Ayush Yadav
* **Date:** 17 October 2026
* **Workbook:** `Food_supply_FAO_2020.xlsm` (Excel Macro-Enabled Workbook)
* **Data Scale:** Real UN FAO per-capita food supply metrics across 216 entities spanning 11,028 rows and 23 columns.
* **Business Objective:** Automate the annual extraction of countries experiencing critical food deficits, eliminating manual multi-step filtering, boundary errors, and repetitive formatting.
* **Integrity Audit:** 0 formula mismatches across all 11,028 independently re-calculated rows.

---

## 📊 Summary of Findings (2017 Early-Warning List)

The screening model evaluates entities using calibrated energy intake thresholds:
* **Severe Deficit:** $< 2,100\text{ kcal/person/day}$ (Priority intervention flag)
* **Warning:** $2,100\text{ to } < 2,500\text{ kcal/person/day}$ (Active watchlist)
* **Adequate:** $\ge 2,500\text{ kcal/person/day}$ (No flag)

### 2017 Action Report Output
The macro isolates the following 8 sovereign nations in Severe Deficit:

| # | Entity | Year | Calories (kcal) | Protein (g) |
|---|---|---|---|---|
| 1 | Afghanistan | 2017 | 2,000 | 54.09 |
| 2 | Bahamas | 2017 | 2,043 | 61.47 |
| 3 | Central African Republic | 2017 | 1,758 | 46.44 |
| 4 | Chad | 2017 | 2,090 | 70.51 |
| 5 | Madagascar | 2017 | 1,903 | 42.94 |
| 6 | North Korea | 2017 | 2,032 | 52.25 |
| 7 | Yemen | 2017 | 2,063 | 53.23 |
| 8 | Zambia | 2017 | 2,013 | 59.09 |



### Key Analytical Insights
1. **Severe Deficit Core:** Central African Republic represents the lowest daily intake at 1,758 kcal (16.3% below the severe threshold), followed by Madagascar at 1,903 kcal.
2. **Long-Term Progress:** Global severe deficit prevalence has dropped substantially over time—from 45.2% (66 of 146 nations) in 1961 down to 4.7% (8 of 172 nations) in 2017.
3. **Compounded Protein Shortfall:** Crisis nations average only 55.00 g/day of protein versus the global sovereign average of 81.30 g/day.
4. **Vulnerable Watchlist:** 33 nations sit in the Warning band just above the line (led closely by Tajikistan at 2,104 kcal and Kenya at 2,124 kcal).

---

## ⚙️ Advanced Excel Functions Implemented

### 1. Entity Disambiguation (`INDEX` / `MATCH`)
Prevents double-counting by looking up entity classifications from `Entity_Types` to separate sovereign nations from regional aggregates (*World*, *Africa*) and historical territories:
```excel
=INDEX(Entity_Types!$B:$B, MATCH(tblFood[[#This Row],[Entity]], Entity_Types!$A:$A, 0))
