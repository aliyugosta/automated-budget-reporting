# ⚙️ Automated Budget Variance Reporting System

> **Business Question:** How do we eliminate 30+ hours of manual monthly reporting while improving accuracy and giving management faster access to budget variance data?

---

## 🗂 Project Overview

Designed and deployed an **Excel VBA macro system** to fully automate the end-to-end budget variance reporting process at Anorma Agro Enterprise Ltd. Prior to this system, finance staff manually compiled and formatted monthly variance reports from ERP exports — a process prone to errors and consuming significant analyst time each month.

The system now auto-generates formatted management reports at the click of a button, integrating data pulled from the ERP system via SQL queries.

| Metric | Value |
|---|---|
| Manual Effort Reduced | 50% |
| Report Cycle | Monthly (automated) |
| Data Source | SAP ERP (via SQL extraction) |
| Output Format | Formatted Excel management report |
| Organisation | Anorma Agro Enterprise Ltd |
| Impact | Faster quarterly decision-making, reduced reporting errors |

---

## 🔧 Tools Used

![Excel VBA](https://img.shields.io/badge/Excel%20VBA-Automation-217346?style=flat&logo=microsoft-excel&logoColor=white)
![SQL](https://img.shields.io/badge/SQL-ERP%20Extraction-4479A1?style=flat&logo=mysql&logoColor=white)
![SAP](https://img.shields.io/badge/SAP-ERP%20System-0FAAFF?style=flat&logo=sap&logoColor=white)

- **Excel VBA** — Macro development, automation logic, report formatting
- **SQL** — Data extraction from SAP ERP system
- **SAP ERP** — Source system for actuals and budget data
- **Advanced Excel** — Budget modelling, variance formulas, conditional formatting

---

## 📸 Project Preview

### Transaction Summary — Multi-Entity Cash Flow Analysis
![Transaction Summary](Screenshot%202.png)

### Healthcare Mass Outreach Budget — 8 Communities
![Mass Outreach Budget Summary](Screenshot%203.png)

---

## 📁 Repository Contents

```
automated-budget-reporting/
│
├── workbook/
│   └── Budget_Variance_Automation.xlsm   # VBA-enabled Excel workbook
│
├── sql/
│   └── erp_data_extraction.sql           # SQL query used for ERP data pull
│
├── docs/
│   └── system_documentation.md           # How to use the system
│
└── README.md
```

---

## ⚙️ How the System Works

### Step 1 — Data Extraction (SQL → ERP)
A parameterised SQL query pulls the current month's actuals from the SAP ERP system:
- Actual spend by department and cost centre
- Budget allocations for the same period
- Prior month actuals for trend comparison

### Step 2 — Data Load (Paste into Master Sheet)
The extracted data is pasted into a designated input sheet. The VBA macro handles all subsequent processing automatically.

### Step 3 — Macro Execution (One Click)
The VBA macro performs the following automatically:
- Calculates variance (Actual vs. Budget) for each line item
- Calculates variance % and flags items exceeding ±10% threshold
- Applies conditional formatting (red/amber/green) to variance columns
- Generates a summary management report on a separate output sheet
- Timestamps the report and clears the previous month's data

### Step 4 — Management Report Output
A clean, formatted report is produced — ready for senior management review without any manual formatting work.

---

## 📊 Sample Variance Logic (VBA Snippet)

```vba
Sub GenerateVarianceReport()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Data Input")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Calculate variance and flag breaches
    For i = 2 To lastRow
        Dim actual As Double
        Dim budget As Double
        Dim variance As Double
        Dim variancePct As Double

        actual = ws.Cells(i, 3).Value
        budget = ws.Cells(i, 4).Value
        variance = actual - budget
        variancePct = IIf(budget <> 0, variance / budget, 0)

        ws.Cells(i, 5).Value = variance
        ws.Cells(i, 6).Value = variancePct

        ' Flag significant variances (>10%)
        If Abs(variancePct) > 0.1 Then
            ws.Cells(i, 6).Interior.Color = RGB(255, 199, 206) ' Red flag
        Else
            ws.Cells(i, 6).Interior.Color = RGB(198, 239, 206) ' Green
        End If
    Next i

    MsgBox "Variance report generated. " & lastRow - 1 & " line items processed.", vbInformation

End Sub
```

---

## 📈 Key Findings & Impact

| Before Automation | After Automation |
|---|---|
| ~6 hours manual compilation per month | ~30 minutes (data paste + one click) |
| Formatting errors common in manual reports | Consistent, error-free output every cycle |
| Report delivered 3–5 days after month end | Report available same day as ERP data |
| No automatic flagging of variance breaches | Automatic red/amber/green flagging |
| Static, manually updated summaries | Dynamic, timestamped management reports |

---

## 💡 Business Recommendations

1. **Extend to quarterly forecasting** — the same VBA logic can be adapted to auto-generate rolling 3-month forecasts by parameterising the date range in the SQL extraction query
2. **Connect directly to ERP via ODBC** — replacing the manual paste step with a live ODBC connection to SAP would make the process fully hands-free
3. **Add email distribution macro** — a VBA Outlook integration can auto-email the finished report to the finance team the moment it's generated, removing the distribution step entirely
4. **Replicate across departments** — the template is modular; deploying it to other cost centres (HR, operations, logistics) would create a company-wide automated reporting layer

---

## 📬 Contact

**Umar Aliyu Gosta** — Financial Data Analyst  
📧 aliyugosta7@gmail.com  
🔗 [LinkedIn](https://linkedin.com/in/YOUR-LINKEDIN-HANDLE)  
🌐 [Portfolio Website](https://aliyugosta.github.io)

> *Open to remote Data Analyst roles with US and European companies.*
