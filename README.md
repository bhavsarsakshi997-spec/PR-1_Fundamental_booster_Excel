# ![Excel Fundamentals Project](https://img.shields.io/badge/Excel%20Fundamentals%20Project-7030A0?style=for-the-badge)

A hands-on Microsoft Excel workbook that demonstrates core-to-intermediate Excel skills across **three realistic datasets**: student grades, sales transactions, and employee records. Each sheet combines clean raw data, calculated columns, and a dedicated "function showcase" area, all with a consistent colour-coding system.

---

## ![Table of Contents](https://img.shields.io/badge/Table%20of%20Contents-7030A0?style=for-the-badge)

1. Project Overview
2. Workbook Structure
3. Colour Coding & Formatting System
4. Sheet 1: Students Grade
5. Sheet 2: Sales Data
6. Sheet 3: Employee Data
7. Functions Index
8. Issues Found & Recommended Fixes
9. Skills Demonstrated
10. How to Use This Workbook

> **Note on formulas:** The formulas below are reconstructed by analysing the sheet outputs, headers and values (every calculation was re-checked by hand against the screenshots). Exact thresholds (e.g. the grade cut-offs) are inferred from the data. If any formula in your file differs slightly, update the snippet here to match.

---

## ![1. Project Overview] <img width="1057" height="680" alt="Screenshot 2026-10-16 111116" src="https://github.com/user-attachments/assets/0d14b33d-a8d3-4fb4-adaf-d1d397000d3d" />


| Item | Details |
|---|---|
| **Tool** | Microsoft Excel (uses modern functions such as `XLOOKUP` / `XMATCH`) |
| **Sheets** | 3: `Students Grade`, `Sales Data`, `Employee Data` |
| **Records** | 20 per sheet (60 total) |
| **Currency** | Indian Rupee (₹) |
| **Goal** | Practise and demonstrate the practical use of logical, lookup, statistical, text, date and reference functions, plus professional formatting |

**Concepts covered:** arithmetic and aggregate formulas, nested `IF`, `IF` with `AND`/`OR`, date arithmetic, rounding, lookups (`VLOOKUP`, `XLOOKUP`, `INDEX`+`MATCH`, `XMATCH`), conditional aggregation (`COUNTIF(S)`, `SUMIF(S)`, `AVERAGEIF(S)`), dynamic references (`INDIRECT`, `OFFSET`) and text formatting (`TEXT`).

---

## ![2. Workbook Structure](https://img.shields.io/badge/2.%20Workbook%20Structure-E67E22?style=for-the-badge)

| Sheet | Purpose | Raw data columns | Calculated columns | Showcase area |
|---|---|---|---|---|
| **Students Grade** | Academic performance analysis | A–F | G–L | `M1:M2` |
| **Sales Data** | Sales and discount logic | A–F | G–J | `A23:J24` and `L2:L3` |
| **Employee Data** | HR and payroll analysis | A–E | F–G | `I1:O8` |

---

## ![3. Colour Coding & Formatting System](https://img.shields.io/badge/3.%20Colour%20Coding%20%26%20Formatting%20System-2980B9?style=for-the-badge)

The workbook uses a **three-colour convention** so a reader can tell what each cell is at a glance:

| Colour | Used for | Meaning |
|---|---|---|
| 🟨 **Cream / light yellow** header | Raw input columns | Data that is typed in (IDs, names, marks, dates, amounts) |
| 🟪 **Purple** header with white text | Calculated columns | Values produced by formulas (Total, Average, Grade, Discount…) |
| 🟩 **Green** header with white text | Function showcase | Stand-alone demonstrations of a single Excel function |
| ⬛ Dark theme | Excel window | Dark mode / dark grid background |

**Other formatting applied**

- Bold, centred, wrapped headers with a taller header row
- Full borders on all data cells; centred alignment
- `₹` currency number format on salary and amount columns (`₹ 7,048.00`)
- Date/time format (`m/d/yyyy h:mm`) on date columns
- Two-decimal formatting for averages
- Green error-indicator triangles on `#NUM!` cells (Excel's built-in error flag)
- Highlighted/selected output cells to show the result of each demo

---

## ![4. Sheet 1: Students Grade](https://img.shields.io/badge/4.%20Sheet%201%3A%20Students%20Grade-C0392B?style=for-the-badge)

![Students Grade sheet output] <img width="1145" height="685" alt="Screenshot 2026-10-16 110716" src="https://github.com/user-attachments/assets/d33de84a-833a-47aa-8234-ab9776d92294" />


### ![4.1 Data dictionary](https://img.shields.io/badge/4.1%20Data%20dictionary-7030A0?style=flat)

| Col | Header | Type | Description |
|---|---|---|---|
| A | Student ID | Input | Unique ID (1001–1020) |
| B | Name | Input | Student 1 … Student 20 |
| C | Math | Input | Marks |
| D | Science | Input | Marks |
| E | English | Input | Marks |
| F | Enrollment Date | Input | Date the student enrolled |
| G | Total Marks | Calc | Sum of the three subjects |
| H | Average | Calc | Mean of the three subjects, 2 decimals |
| I | Grade | Calc | Letter grade from the average |
| J | Result | Calc | Pass / Fail |
| K | Age | Calc | Years since enrollment |
| L | IF with AND | Calc | Eligible / Not Eligible |
| M | Students With Math>60 | Calc | Count of students scoring above 60 in Math |

### ![4.2 Formulas explained](https://img.shields.io/badge/4.2%20Formulas%20explained-7030A0?style=flat)

**Total Marks (G)**
```excel
=SUM(C2:E2)
```
Example: Student 1 → 61 + 48 + 76 = **185**.

**Average (H)**
```excel
=ROUND(AVERAGE(C2:E2),2)
```
Example: 185 / 3 = **61.67**.

**Grade (I)**: nested `IF` on the average
```excel
=IF(H2>=80,"A",IF(H2>=60,"B","C"))
```
Observed behaviour: averages of 84.33, 95 → **A**; 61.67–78.33 → **B**; 55–59 → **C**. The cut-offs (80 / 60) are consistent with every row.

**Result (J)**: every student shows **Pass** (all subject marks are 40 or higher, so a pass-mark rule is met).
```excel
=IF(H2>=40,"Pass","Fail")
```

**Age (K)**: whole years since enrollment
```excel
=DATEDIF(F2,TODAY(),"Y")
```
Example: 1/1/2018 → **8** years (as of Oct 2026). Students enrolled within the last year show **0**.

**IF with AND (L)**: multi-condition eligibility test
```excel
=IF(AND(C2>=80, <second condition>),"Eligible","Not Eligible")
```
Observed behaviour: every student with **Math ≥ 80** is *Eligible* (80, 82, 83, 85, 90, 93, 94, 97, 98) and every student below 80 is *Not Eligible* (even 78). This demonstrates how `AND()` requires **all** conditions to be true.

**Students with Math > 60 (M2)**
```excel
=COUNTIF(C2:C21,">60")
```
Result **19** (only Student 9 scored 41).

### ![4.3 Key insights from the output](https://img.shields.io/badge/4.3%20Key%20insights%20from%20the%20output-7030A0?style=flat)

- **Top performer:** Student 13 (Total 285, Average 95, Grade A)
- **Grade A:** 2 students (Student 12 and Student 13)
- **Grade C:** 4 students (Students 4, 9, 14, 17)
- **Pass rate:** 100%
- **Highest Math mark:** 98 (Students 10, 13, 16); **lowest:** 41 (Student 9)

---

## ![5. Sheet 2: Sales Data](https://img.shields.io/badge/5.%20Sheet%202%3A%20Sales%20Data-16A085?style=for-the-badge)

![Sales Data sheet output] <img width="1057" height="680" alt="Screenshot 2026-10-16 111116" src="https://github.com/user-attachments/assets/18cb7632-b76e-44f3-8c09-f04821972fbd" />


### ![5.1 Data dictionary](https://img.shields.io/badge/5.1%20Data%20dictionary-00A86B?style=flat)

| Col | Header | Type | Description |
|---|---|---|---|
| A | Sales ID | Input | 2001–2020 |
| B | Product | Input | Keyboard, Printer, Mouse, Monitor, Laptop |
| C | Region | Input | North, South, East, West |
| D | Salesperson | Input | Person 1 … Person 20 |
| E | Amount | Input | Sale value in ₹ |
| F | Date | Input | Sale date (2023) |
| G | Discount Amount | Calc | Discount slab: No Discount / 5% / 10% |
| H | Discount Eligible | Calc | Eligible / Not eligible |
| I | Celling Amount | Calc | Amount rounded **up** to the next 100 |
| J | Floor Amount | Calc | Amount rounded **down** to the previous 100 |

### ![5.2 Formulas explained](https://img.shields.io/badge/5.2%20Formulas%20explained-00A86B?style=flat)

**Discount slab (G)**: nested `IF`
```excel
=IF(E2>=25000,"10%Discount",IF(E2>=10000,"5%Discount","No Discount"))
```
Observed behaviour: amounts under about ₹10,000 → *No Discount*; about ₹15,000–₹24,000 → *5%*; ₹30,000 and above → *10%*. (The exact upper cut-off lies between ₹23,672 and ₹30,436.)

**Discount Eligible (H)**: combines two conditions with `OR`
```excel
=IF(OR(C2="South",E2>=25000),"Eligible","Not eligible")
```
Observed behaviour: **all South-region sales are Eligible** (even small ones like ₹7,048), while sales from other regions are Eligible only for large amounts.

**Ceiling Amount (I)**
```excel
=CEILING(E2,100)      // 7,048 → 7,100
```

**Floor Amount (J)**
```excel
=FLOOR(E2,100)        // 7,048 → 7,000
```

### ![5.3 Function showcase (row 24)](https://img.shields.io/badge/5.3%20Function%20showcase%20%28row%2024%29-00A86B?style=flat)

| Header | Output | Formula (reconstructed) | What it demonstrates |
|---|---|---|---|
| XMATCH – Position of Product | `1` | `=XMATCH("Keyboard",B2:B21)` | Returns the position of the first match in a list |
| INDIRECT Function | `7048` | `=INDIRECT("E2")` | Reads a cell from a **text** address |
| OFFSET Function | `17428` | `=OFFSET(E2,1,0)` | Moves 1 row down from E2 |
| ROUND | `7048` | `=ROUND(E2,0)` | Rounds to a set number of digits |
| F3 − F2 | `15` | `=F3-F2` | Date subtraction → **15 days** between sales |
| INDIRECT Function (2) | `7048` | `=INDIRECT(L2)` | `INDIRECT` driven by the address typed in cell `L2` ("E2") |
| VLOOKUP Function | `₹ 17,428.00` | `=VLOOKUP(2002,A2:E21,5,FALSE)` | Finds the amount for Sales ID 2002 |
| TEXT | `01-Jan-2023` | `=TEXT(F2,"dd-mmm-yyyy")` | Converts a date into formatted text |
| INDIRECT (3) | `2001` | `=INDIRECT("A2")` | Returns the first Sales ID |

### ![5.4 Key insights from the output](https://img.shields.io/badge/5.4%20Key%20insights%20from%20the%20output-00A86B?style=flat)

- **Highest sale:** ₹46,429 (Sales ID 2008, Keyboard, East)
- **Lowest sale:** ₹5,264 (Sales ID 2004, Printer, East)
- **10% discount tier** is the largest group, with 10 of 20 transactions
- **No Discount:** 4 transactions (all below ₹8,000)
- Keyboards and Monitors generate the highest-value sales

---

## ![6. Sheet 3: Employee Data](https://img.shields.io/badge/6.%20Sheet%203%3A%20Employee%20Data-8E44AD?style=for-the-badge)

![Employee Data sheet output] <img width="1917" height="690" alt="Screenshot 2026-10-16 111433" src="https://github.com/user-attachments/assets/615f4da9-ff8c-45d8-9b96-0977a639bef1" />



### ![6.1 Data dictionary](https://img.shields.io/badge/6.1%20Data%20dictionary-E67E22?style=flat)

| Col | Header | Type | Description |
|---|---|---|---|
| A | Employee ID | Input | 3001–3020 |
| B | Name | Input | Employee 1 … Employee 20 |
| C | Department | Input | Finance, HR, IT, Marketing |
| D | Salary | Input | Annual salary in ₹ |
| E | Joining Date | Input | Date of joining |
| F | Bonus | Calc | 10% of salary |
| G | Salary + bonus | Calc | Total compensation |

### ![6.2 Core formulas](https://img.shields.io/badge/6.2%20Core%20formulas-E67E22?style=flat)

**Bonus (F)**
```excel
=D2*10%        // 86,766 × 10% = 8,676.60
```

**Salary + Bonus (G)**
```excel
=D2+F2         // 86,766 + 8,676.60 = 95,442.60
```

### ![6.3 Function showcase (right-hand panel)](https://img.shields.io/badge/6.3%20Function%20showcase%20%28right--hand%20panel%29-E67E22?style=flat)

| Panel title | Output | Formula (reconstructed) | What it demonstrates |
|---|---|---|---|
| **XLOOKUP Function** | `31943` | `=XLOOKUP(3002,A2:A21,D2:D21)` | Modern lookup: returns the salary for ID 3002 |
| XLOOKUP (date) | `42005` | `=XLOOKUP(3001,A2:A21,E2:E21)` | Returns a joining date; shows as a serial number (42005 = 1 Jan 2015) because the cell is not date-formatted |
| **VLOOKUP Function** | `Employee 1` | `=VLOOKUP(3001,A2:B21,2,FALSE)` | Finds the employee name |
| **INDEX Function** | `86766` | `=INDEX(D2:D21,1)` | Returns the item at a given position |
| **D2+H2** | `86766` | `=D2+H2` | Cell arithmetic (H2 is blank, so the result equals the salary) |
| **SUMIFC (SUMIF)** | `417713` | `=SUMIF(C2:C21,"HR",D2:D21)` | Total salary of the HR department |
| **COUNTIFC (COUNTIF)** | `6` | `=COUNTIF(C2:C21,"HR")` | Number of HR employees |
| **COUNTIFC IT** | `5` | `=COUNTIF(C2:C21,"IT")` | Number of IT employees |
| **AVERAGEIF** | `80952.33` | `=AVERAGEIF(C2:C21,"Finance",D2:D21)` | Average Finance salary (485,714 ÷ 6) |
| **XLOOKUP Fuc** | `Finance` | `=XLOOKUP(3001,A2:A21,C2:C21)` | Returns the department of Employee 1 |
| **INDEX + MATCH SALARY** | `86766` | `=INDEX(D2:D21,MATCH(3001,A2:A21,0))` | Classic two-function lookup, returning salary |
| **INDEX + MATCH DEPARTMENT** | `Finance` | `=INDEX(C2:C21,MATCH(3001,A2:A21,0))` | Same technique, returning department |
| **HR DEPT_TOTAL SALARY** | `417713` | `=SUMIFS(D2:D21,C2:C21,"HR")` | Conditional sum using `SUMIFS` |
| **DEPARTMENT TOTAL SALARY** | `485714` | `=SUMIF(C2:C21,"Finance",D2:D21)` | Total Finance salary |
| **DATEIF Function** | `11` | `=DATEDIF(E2,TODAY(),"Y")` | Years of service (Jan 2015 → Oct 2026 = 11) |

### ![6.4 Department summary (derived from the data)](https://img.shields.io/badge/6.4%20Department%20summary%20%28derived%20from%20the%20data%29-E67E22?style=flat)

| Department | Employees | Total Salary (₹) | Average Salary (₹) |
|---|---|---|---|
| Finance | 6 | 4,85,714 | 80,952.33 |
| HR | 6 | 4,17,713 | 69,618.83 |
| IT | 5 | 4,31,204 | 86,240.80 |
| Marketing | 3 | 2,52,167 | 84,055.67 |

**Insights:** IT has the highest average salary; the highest individual salary is ₹1,16,284 (Employee 18, IT); the lowest is ₹31,943 (Employee 2, Finance).

---

## ![7. Functions Index](https://img.shields.io/badge/7.%20Functions%20Index-D35400?style=for-the-badge)

| Category | Functions used |
|---|---|
| **Math & Aggregation** | `SUM`, `AVERAGE`, `ROUND`, `CEILING`, `FLOOR` |
| **Logical** | `IF` (nested), `AND`, `OR` |
| **Lookup & Reference** | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`, `MATCH`, `INDIRECT`, `OFFSET` |
| **Conditional Statistics** | `COUNTIF`, `SUMIF`, `SUMIFS`, `AVERAGEIF` |
| **Date & Time** | `DATEDIF`, `TODAY`, date subtraction |
| **Text** | `TEXT` |

---

## ![8. Issues Found & Recommended Fixes](https://img.shields.io/badge/8.%20Issues%20Found%20%26%20Recommended%20Fixes-2C3E50?style=for-the-badge)

Reviewing the output screenshots turned up the following items worth correcting before presenting or publishing the workbook.

| # | Sheet / Cell | Issue | Recommended fix |
|---|---|---|---|
| 1 | Students Grade `K20:K21` | **`#NUM!` error.** `DATEDIF` fails when the enrollment date is in the future (11/15/2026, 5/14/2027). | `=IF(F20>TODAY(),"Not enrolled yet",DATEDIF(F20,TODAY(),"Y"))`, or wrap in `IFERROR` |
| 2 | Students Grade `L` | Typo **"Egiible"** appears in several rows (the formula text is spelled wrong). | Correct the string to `"Eligible"` |
| 3 | Sales Data `L3` | Shows `6/23/1905`: the value `2001` is formatted as a **date**. | Set cell format to **General** |
| 4 | Employee Data `K2` | Shows `10/15/1911`, and `I3` shows `42005`: numbers formatted as dates, and a date formatted as a number. | Apply the correct Date / General format |
| 5 | Employee Data `E19:E21` | Joining dates are in the **future** (2026–2028), which is unrealistic for current employees. | Correct the dataset |
| 6 | Sales Data headers | Spelling: *"Celling Amount"*, *"INDIRACT Function"*, *"Not eligible/Eligible"* casing inconsistent. | Rename to *Ceiling Amount*, *INDIRECT Function*; use one casing |
| 7 | Employee Data headers | Spelling: *"SUMIFC Fuction"*, *"COUNTIFC Fuction"*, *"AVARAGEIFC"*, *"DATEIFC"*, *"XLOOKUP Fuc"*, *"DEPARMENT"*. | Use *SUMIF Function*, *COUNTIF Function*, *AVERAGEIF Function*, *DATEDIF Function*, *XLOOKUP Function*, *DEPARTMENT* |
| 8 | Employee Data `K4` | `D2+H2` references an empty cell (H2). | Change to `D2+F2` (or `=G2`) |
| 9 | All sheets | Hard-coded thresholds (80, 60, 25000, 10%) inside formulas. | Move them to a small **assumptions table** and reference with absolute refs (`$`) |
| 10 | All sheets | Plain ranges. | Convert data to **Excel Tables** (`Ctrl+T`) for structured references and auto-expanding formulas |

---

## ![9. Skills Demonstrated](https://img.shields.io/badge/9.%20Skills%20Demonstrated-E91E63?style=for-the-badge)

- Structuring data into clean, analysis-ready tables
- Building calculated columns with arithmetic, logical and date functions
- Writing nested and multi-condition logic (`IF`, `AND`, `OR`)
- Performing exact-match lookups with both legacy and modern functions
- Summarising data conditionally by category (`SUMIF`, `COUNTIF`, `AVERAGEIF`)
- Working with dynamic references (`INDIRECT`, `OFFSET`)
- Professional presentation: consistent colour coding, borders, number formats and currency formatting
- Spotting and diagnosing formula errors (`#NUM!`, wrong cell formats)

---

## ![10. How to Use This Workbook](https://img.shields.io/badge/10.%20How%20to%20Use%20This%20Workbook-27AE60?style=for-the-badge)

1. Open the `.xlsx` file in **Excel 2021 / Microsoft 365** (required for `XLOOKUP` and `XMATCH`).
2. Switch between the three tabs at the bottom: **Students Grade**, **Sales Data**, **Employee Data**.
3. Edit any **cream** (input) cell and watch the **purple** (calculated) cells update automatically.
4. Change the lookup keys (e.g. an Employee ID or Sales ID) inside the **green** showcase formulas to test each function.
5. Press `Ctrl + ~` to toggle formula view and inspect every formula.

---

### ![Repository Structure](https://img.shields.io/badge/Repository%20Structure-2980B9?style=flat)

```
excel-fundamentals-project/
├── README.md
├── ExcelFundamentals.xlsx      <- add your workbook here
└── images/
    ├── students-grade.png
    ├── sales-data.png
    └── employee-data.png
```

### ![Video link] https://drive.google.com/file/d/18ARk4pundYd4_cnNk7-p-LMekeVAnUl4/view?usp=drive_link

