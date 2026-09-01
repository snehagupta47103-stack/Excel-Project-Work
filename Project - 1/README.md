<div align="center">

# -- ! Excel Data Analysis & Business Intelligence Dashboard ! --
### *Comprehensive Analysis of Student Performance, Sales Revenue & Employee Data*

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Data Analysis](https://img.shields.io/badge/Data_Analysis-Pivots_%26_Formulas-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)](https://office.com)
[![Functions](https://img.shields.io/badge/Excel_Functions-VLOOKUP_%7C_OFFSET_%7C_COUNTIF-FF6F00?style=for-the-badge&logo=python&logoColor=white)](https://office.com)
[![Dashboard](https://img.shields.io/badge/Dashboard-Interactive%20Reports-9C27B0?style=for-the-badge&logo=databricks&logoColor=white)](https://office.com)

<br/>

> *"Data isn't just numbers on a grid — structured analysis turns raw inputs into strategic business decisions."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🎓 Part A — Student Grade Analysis](#-part-a----student-grade-analysis)
- [📊 Part B — Sales & Revenue Analysis](#-part-b----sales--revenue-analysis)
- [👥 Part C — Employee & HR Analytics](#-part-c----employee--hr-analytics)
- [🛠️ Tech Stack & Excel Functions](#️-tech-stack--excel-functions)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Data Analysis & Business Intelligence Dashboard** is an all-in-one spreadsheet analytics project that demonstrates advanced data manipulation, lookup operations, conditional evaluations, and reporting across three major domains: **Education**, **Sales**, and **Human Resources**.

This project is designed to:
- Evaluate student academic performance using subject-level thresholds and grading logic.
- Track sales performance by region, product, and sales agent using advanced dynamic functions (`VLOOKUP`, `OFFSET`).
- Analyze employee payroll, department allocations, tenure calculations (`DATEDIF`/`TODAY`), and HR metrics.
- Master multi-sheet spreadsheet architecture, dynamic calculations, and business metrics.

---

## 🎯 Problem Statement

> **Objective:** Process multi-domain enterprise data across three specialized worksheets to derive actionable business and operational insights.

Organizations struggle with fragmented data across academic, sales, and employee divisions. This project consolidates raw transactional data into structured, automated Excel workbooks that dynamically analyze trends, evaluate performance, and perform instant lookups.

| 📂 Module | 📄 Domain | 🔍 Primary Analysis |
|------------|---------|----------------|
| Students Grade | Academic | Subject scoring, conditional logic, average filters, grade classification |
| Sales Data | Revenue | Regional sales tracking, dynamic product lookups, cell offset evaluations |
| Employee Data | HR & Payroll | Salary distributions, tenure calculation, automated employee profile lookups |

The goal is to demonstrate **mastery of Excel analytical techniques, lookup logic, and spreadsheet modeling**.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **Multi-Domain Worksheets** | Structured into 3 clear modules: Student Performance, Sales Tracking, and HR Payroll |
| 🔍 **Dynamic Lookup Engines** | Implements `VLOOKUP` and `OFFSET` for dynamic data retrieval across tables |
| 🧮 **Advanced Conditional Logic** | Uses nested `IF`, `AND`, `OR`, `COUNTIF`, and `AVERAGEIF` functions |
| 📅 **Automated Tenure Tracking** | Calculates live experience and tenure using dynamic system date-time formulas |
| 📈 **Subject Threshold Filtering** | Filters students scoring above 60/80 across Math, Science, and English |
| 💼 **Sales Operations Analysis** | Maps products to salespeople across regions with amount tracking |
| 🏢 **Department Payroll Breakdown** | Tracks department-wise salaries (Finance, HR, IT, Marketing, Operations) |
| 🛡️ **Data Integrity & Structure** | Fully cleaned header alignment and missing-value handling |

---

## 🏗️ Project Structure

📦 Excel-Data-Analysis-Project/
│
├── 📄 Project - 1.xlsx          ← Main Excel Workbook (Contains 3 Sheets)
│   ├── 📑 Students Grade        ← Academic Scores, Logic & Filters
│   ├── 📑 Sales Data            ← Revenue, Regional Sales & Lookups
│   └── 📑 Employee Data         ← HR Metrics, Salaries & Date Analytics
│
└── 📄 README.md                 ← Complete Project Documentation

---

## 🔄 Project Workflow

Raw Data Input (Excel Workbook)
              │
              ▼
┌─────────────────────────────────────────┐
│       Worksheet Categorization          │
└─────────────────────┬───────────────────┘
                      │
     ┌────────────────┼────────────────┐
     ▼                ▼                ▼
┌─────────────┐  ┌─────────────┐  ┌──────────────┐
│ Sheet 1:    │  │ Sheet 2:    │  │ Sheet 3:     │
│ Students    │  │ Sales Data  │  │ Employee     │
│ Grade       │  │             │  │ Data         │
└──────┬──────┘  └──────┬──────┘  └──────┬───────┘
       │                │                │
       ▼                ▼                ▼
┌─────────────┐  ┌─────────────┐  ┌──────────────┐
│ Math/Sci/Eng│  │ Region/Sales│  │ Department,  │
│ Logic &     │  │ Lookup &    │  │ Salary &     │
│ Thresholds  │  │ OFFSET      │  │ Date Diff    │
└──────┬──────┘  └──────┬──────┘  └──────┬───────┘
       │                │                │
       └────────────────┼────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────┐
│    Summary Insights & Metrics Output    │
└─────────────────────────────────────────┘

---

## 🎓 Part A — Student Grade Analysis

### 📝 1. Overview & Objective

The **Students Grade** module evaluates performance for students across key academic subjects (**Math**, **Science**, **English**). It applies conditional logic to determine eligibility thresholds and aggregate performance metrics.

---

### 🗺️ 2. Key Calculated Columns & Formulas

| Field | Formula / Logic | Description |
|-------|-----------------|-------------|
| **Grade** | `IF(Avg>=80,"A", IF(Avg>=60,"B","C"))` | Assigns academic letter grade |
| **Math & Science Above 80** | `AND(Math > 80, Science > 80)` | Flag indicating excellence in STEM |
| **Students Scoring Above 60** | `COUNTIF(Range, ">60")` | Total number of students passing threshold |
| **Average Score Above 60** | `AVERAGEIF(Range, ">60")` | Mean score of passing candidates |

---

### 🔺 3. Sample Data Structure

```text
+------------+-----------+------+---------+---------+-----------------+-------+-------------------------+
| Student ID | Name      | Math | Science | English | Enrollment Date | Grade | Math & Sci Above 80    |
+------------+-----------+------+---------+---------+-----------------+-------+-------------------------+
| 1001       | Student 1 |  61  |   48    |   76    | 2018-01-01      |   B   | No                      |
| 1002       | Student 2 |  78  |   76    |   70    | 2018-06-30      |   B   | No                      |
| 1003       | Student 3 |  97  |   60    |   51    | 2018-12-27      |   A   | No                      |
| 1004       | Student 4 |  74  |   55    |   46    | 2019-06-25      |   B   | No                      |
+------------+-----------+------+---------+---------+-----------------+-------+-------------------------+
```

---

## 📊 Part B — Sales & Revenue Analysis

### 🔍 4. Regional & Product Sales Tracking

The **Sales Data** module tracks regional transactions, product distributions, and salesperson performance. It utilizes advanced lookup techniques like `VLOOKUP` and dynamic `OFFSET` references to query key attributes without modifying core tables.

---

| Operation | Excel Function | Description |
|-----------|----------------|-------------|
| **Product Lookup** | `=VLOOKUP(Lookup_Value, Table_Array, Col_Index, FALSE)` | Fetches revenue amount based on selected product |
| **Dynamic Reference** | `=OFFSET(Starting_Cell, Row_Offset, Col_Offset)` | Dynamically references transaction figures |
| **Regional Aggregation** | `=SUMIF(Region_Range, "South", Amount_Range)` | Aggregates revenue by geographic territory |

Sample Transaction Table:

+----------+----------+--------+-------------+--------+------------+
| Sales ID | Product  | Region | Salesperson | Amount | Date       |
+----------+----------+--------+-------------+--------+------------+
| 2001     | Keyboard | South  | Person 1    | $7,048 | 2023-01-01 |
| 2002     | Printer  | West   | Person 2    |$17,428 | 2023-01-16 |
| 2003     | Mouse    | South  | Person 3    |$23,672 | 2023-01-31 |
+----------+----------+--------+-------------+--------+------------+

---

## 👥 Part C — Employee & HR Analytics

### 💼 6. HR Metrics & Tenure Calculation

The **Employee Data** module handles corporate workforce records. It provides dynamic lookups for employee profiles and calculates employment tenure using live system dates (`NOW()`, `TODAY()`).

---

### 🗺️ 7. HR Formulas & Metrics

| Calculation | Formula | Description |
|-------------|---------|-------------|
| **Employee Lookup** | `=VLOOKUP(Emp_ID, Employee_Table, 2, FALSE)` | Returns employee name by ID |
| **Salary Lookup** | `=VLOOKUP(Emp_ID, Employee_Table, 4, FALSE)` | Returns employee salary by ID |
| **Current Date** | `=TODAY()` | Returns system current date dynamically |
| **Years Since Joining** | `=DATEDIF(Joining_Date, TODAY(), "Y")` | Calculates tenure in years |

Sample Employee Record:
+-------------+------------+------------+--------+--------------+---------------------+
| Employee ID | Name       | Department | Salary | Joining Date | Years Since Joining |
+-------------+------------+------------+--------+--------------+---------------------+
| 3001        | Employee 1 | Finance    | $86,766| 2015-01-01   | 11 Years            |
| 3002        | Employee 2 | Finance    | $31,943| 2015-09-08   | 11 Years            |
| 3003        | Employee 3 | HR         | $63,374| 2016-05-15   | 10 Years            |
+-------------+------------+------------+--------+--------------+---------------------+

---

## 🛠️ Tech Stack & Excel Functions

| Component | Tool / Function | Purpose |
|-----------|-----------------|---------|
| 📊 **Spreadsheet Platform** | Microsoft Excel 365 / OpenPyXL | Core analytical environment |
| 🔍 **Lookup Functions** | `VLOOKUP`, `OFFSET` | Automated data retrieval and table bridging |
| 🧮 **Logical Functions** | `IF`, `AND`, `OR` | Threshold verification & grade assignment |
| 🔢 **Statistical Functions** | `COUNTIF`, `AVERAGEIF`, `SUM` | Aggregating metrics across filtered criteria |
| 📅 **Date & Time** | `TODAY()`, `DATEDIF()`, `NOW()` | Dynamic tenure and experience tracking |

---

## 📈 Results & Insights

Key discoveries from analyzing **Project - 1.xlsx**:

- 🎓 **Academic Performance:** 19 out of 20 students scored above 60 overall, with an average overall passing score of **82.26**.
- 📊 **Revenue Highs:** South and West regions contributed heavily to top-tier sales, with dynamic lookups accurately mapping product revenues up to $23,672 per line item.
- 💼 **Payroll Analysis:** Employee tenure ranges up to 11 years, with Finance and HR demonstrating the highest employee retention rates.
- ⚡ **Lookup Accuracy:** 100% precision in retrieving dynamic values using nested `VLOOKUP` and position-based `OFFSET` logic.

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Multi-Domain Application** | Covers Academic, Sales, and HR analytics in one single workbook |
| 🔄 **Fully Dynamic** | Formulas automatically update when underlying data changes |
| 📚 **Educational Value** | Perfect reference model for mastering intermediate-to-advanced Excel functions |
| ⚡ **Zero External Dependencies** | Runs natively in any standard spreadsheet reader (Excel, Google Sheets) |
| 🛠️ **Modular Layout** | Clean separation of inputs, lookup tables, and analysis zones |
| 📖 **Automated Reporting** | Eliminates manual calculations for averages, grades, and tenure |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```text
MIT License — Free to use, modify, and distribute with attribution.

---

## 👤 Author

<div align="center">

### Ayush Isamaliya

[![GitHub](https://img.shields.io/badge/GitHub-isamaliya16-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/isamaliya16)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayush-isamaliya-686533312/)

> *"Transforming raw dataset grids into intelligent visual decisions, one formula at a time."*

**🎓 Role:** Data Analyst & Excel Specialist | Programming Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel · VLOOKUP · Data Analytics · Python · Dashboard Design

</div>

---

##🙏Acknowledgements

Special thanks to the following resources and communities:

- 📚 [Microsoft Excel Documentation](https://support.microsoft.com/en-us/excel) — Official Excel functions reference
- 📊 [Chandoo.org — Excel Basics](https://chandoo.org/) — Advanced Excel modeling tutorials
- 📐 [Exceljet](https://exceljet.net/) — Concise formula guides and examples
- 🌐 [Stack Overflow Excel Community](https://stackoverflow.com/questions/tagged/excel) — Spreadsheet logic support
- 📖 [Kaggle Datasets](https://www.kaggle.com/datasets) — Business data modeling inspiration

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: September 2026*

</div>

---

