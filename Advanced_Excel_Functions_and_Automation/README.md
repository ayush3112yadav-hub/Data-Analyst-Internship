# Week 3: Advanced Excel Functions & VBA Automation

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
```

### 2. Multi-Tier Status Grading (Nested IF)
Enforces strict nutritional thresholds across all 11,028 records:
```excel
=IF(D2 < 2100, "Severe Deficit", IF(D2 < 2500, "Warning", "Adequate"))
```

### 3. Conditional Array Averaging (SUMPRODUCT)
Calculates the exact average protein supply for the 2017 Severe Deficit cohort without creating intermediate helper columns:
```excel
=SUMPRODUCT(--(C2:C11029=2017), --(B2:B11029="Country"), --(W2:W11029="Severe Deficit"), F2:F11029) 
 / SUMPRODUCT(--(C2:C11029=2017), --(B2:B11029="Country"), --(W2:W11029="Severe Deficit"))
```

### 4. Scoped Extremum Lookup (INDEX / MATCH / MAXIFS)
Resolves the highest caloric value specifically for sovereign countries in 2017 (Belgium at 3,768 kcal) rather than picking historical legacy entities across all years:
```excel
=INDEX(A2:A11029, MATCH(1, INDEX((C2:C11029=2017)*(B2:B11029="Country")*(D2:D11029=MAXIFS(D2:D11029, C2:C11029, 2017, B2:B11029, "Country")), 0), 0))
```

### 5. Calculated Balance Validation (Structured References)
Audits individual meat commodity breakdowns against overall reported meat values using dynamic structured table references:
```excel
=IF(ABS(tblFood[[#This Row],[Meat_Check]]) <= 0.2, "Pass", "Check")
```

## 💻 VBA Automation (GenerateActionReport)
The VBA procedure scans Clean_Data, filters matching rows, initializes or resets the target sheet, styles headers, applies light-red urgency highlighting, and notifies the user with a summary dialog.
```VBA
Sub GenerateActionReport()
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, i As Long, reportRow As Long
    
    ' Set source sheet
    Set wsData = ThisWorkbook.Sheets("Clean_Data")
    
    ' Initialize or clear target sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Action_Required")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=wsData)
        wsReport.Name = "Action_Required"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Format Headers
    wsReport.Range("A1:D1").Value = Array("Entity", "Year", "Calories (kcal)", "Protein (g)")
    wsReport.Range("A1:D1").Font.Bold = True
    wsReport.Range("A1:D1").Interior.Color = RGB(200, 200, 200)
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    reportRow = 2
    
    ' Extract 2017 Severe Deficit Countries
    Application.ScreenUpdating = False
    For i = 2 To lastRow
        ' Col 3 = Year, Col 2 = Type, Col 23 = Nutritional Status (W)
        If wsData.Cells(i, 3).Value = 2017 And _
           wsData.Cells(i, 2).Value = "Country" And _
           wsData.Cells(i, 23).Value = "Severe Deficit" Then
            
            ' Transfer Entity, Year, Calories, Protein
            wsReport.Cells(reportRow, 1).Value = wsData.Cells(i, 1).Value
            wsReport.Cells(reportRow, 2).Value = wsData.Cells(i, 3).Value
            wsReport.Cells(reportRow, 3).Value = wsData.Cells(i, 4).Value
            wsReport.Cells(reportRow, 4).Value = wsData.Cells(i, 6).Value
            
            ' Apply urgency highlighting
            wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, 4)).Interior.Color = RGB(255, 200, 200)
            reportRow = reportRow + 1
        End If
    Next i
    
    wsReport.Columns("A:D").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Action report generated successfully. Found " & (reportRow - 2) & " crisis zones.", vbInformation
End Sub
```

## 🧪 Quality Assurance & Test Matrix
To guarantee reporting reliability, the model underwent four layers of testing: boundary tests, recalculations, reconciliation audits, and execution tests.

| # | Test | Expected result | Expected value | Observed | Result |
|---|---|---|---|---|:---:|
| **M1** | Run the macro once | Message box and a filled report sheet | 8 data rows; "Found 8 crisis zones" | 8 rows on Action_Required in the submitted workbook | PASS |
| **M2** | Run it a second time | The sheet is cleared and rebuilt, not appended to | Still 8 rows | Still 8 rows; previous contents cleared without duplicate rows appended | PASS |
| **M3** | Delete Action_Required, then run | The sheet is created again | 8 rows | Sheet created automatically and populated with 8 rows | PASS |
| **M4** | Change 2100 to 2200 in the grading formula, then run | More countries are flagged | 15 rows | 15 rows generated; message box displayed "Found 15 crisis zones" | PASS |
| **M5** | Change 2100 to 2000, then run | Fewer countries are flagged | 2 rows (Central African Republic, Madagascar) | 2 rows generated (Central African Republic and Madagascar); "Found 2 crisis zones" | PASS |
| **M6** | Inspect the output Type | No regions or former states appear | 0 non-country rows | 0 non-country rows; all 8 records are confirmed Country type | PASS |
| **M7** | Check the first and last rows | Same order as the data | Afghanistan first, Zambia last | First row is Afghanistan (2,000 kcal), last row is Zambia (2,013 kcal) | PASS |

### Mathematical Reconciliation Checks
1: R1 (Full Table Count): $\text{Adequate (5,994)} + \text{Warning (2,980)} + \text{Severe Deficit (2,054)} = 11,028\text{ rows}$[cite: 2].
2: R2 (2017 Sovereign Total): $131\text{ (Adequate)} + 33\text{ (Warning)} + 8\text{ (Severe Deficit)} = 172\text{ sovereign nations}$[cite: 2].
3: R3 (Protein Ratio): Aggregate protein sum ($440.02\text{ g}$) divided by $8\text{ nations} = 55.0025\text{ g/day}$ (matching cell Y2)[cite: 2].

## 📁 Repository Structure



## 🚀 How to Run the Macro
1: Open Food_supply_FAO_2020.xlsm using desktop Microsoft Excel on Windows or macOS[cite: 2].
2: Click Enable Content in the yellow Trust Center security bar if prompted[cite: 2].
3: Press Alt + F8, select GenerateActionReport, and click Run[cite: 2].
4: Confirm the completion message ("Found 8 crisis zones") and view the freshly formatted table on the Action_Required sheet[cite: 2].

<FollowUp>
Would you like me to tailor any specific section for your repository (such as adding installation badges or contribution guidelines)?
</FollowUp>
