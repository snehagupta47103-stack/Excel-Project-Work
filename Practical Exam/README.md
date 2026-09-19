<div align="center">

# 📊 Excel Sales Performance & Interactive Dashboard

### *Sales Data Analysis, Advanced Excel Formulas & Interactive Dashboard*

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-Data%20Analysis-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Excel-blue?style=for-the-badge)
![Dashboard](https://img.shields.io/badge/Interactive-Dashboard-purple?style=for-the-badge)
![BBA](https://img.shields.io/badge/BBA-Student-orange?style=for-the-badge)

> **“Turning raw sales data into meaningful insights through Excel.”**

</div>

---

# 📑 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [📁 Project Structure](#-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧹 Part A — Data Cleaning & Preparation](#-part-a--data-cleaning--preparation)
- [📊 Part B — Pivot Table Analysis](#-part-b--pivot-table-analysis)
- [🧮 Part C — Advanced Excel Formulas](#-part-c--advanced-excel-formulas)
- [📈 Part D — Data Visualization](#-part-d--data-visualization)
- [📋 Part E — Interactive Dashboard](#-part-e--interactive-dashboard)
- [📝 Part F — Documentation](#-part-f--documentation)
- [🛠️ Tech Stack](#️-tech-stack)
- [📊 Results & Learning Outcomes](#-results--learning-outcomes)
- [⭐ Advantages](#-advantages)
- [📜 License](#-license)
- [👩‍💻 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

# 📌 Overview

The **Excel Sales Performance & Interactive Dashboard** project is a practical Excel-based data analysis project designed to analyze sales performance using a structured dataset of **228 sales records across 13 columns**.

The project demonstrates the complete workflow of transforming raw sales data into meaningful business insights using **Excel Tables, Pivot Tables, Advanced Formulas, Conditional Formatting, Charts, KPI Cards, Slicers, Timelines, and an Interactive Dashboard**.

The workbook is organized into four main sheets:

- `Dataset`
- `Pivot Tables`
- `Main Task`
- `Dashboard`

The project focuses on practical applications of Microsoft Excel for **data cleaning, analysis, visualization, and business reporting**.

---

# 🎯 Problem Statement

Businesses generate large amounts of sales data that need to be organized, analyzed, and presented clearly for effective decision-making.

The objective of this project is to:

- Clean and prepare a sales dataset.
- Organize the dataset into a structured Excel Table.
- Analyze sales performance using Pivot Tables.
- Apply advanced Excel formulas for data lookup and classification.
- Create meaningful charts and visualizations.
- Build an interactive sales dashboard.
- Present important KPIs in an easy-to-understand format.
- Document the complete Excel analysis workflow.

---

# ✨ Key Features

### 📂 Dataset Management

- 228 sales records.
- 13 data columns.
- Structured Excel Table named `SalesData`.
- Proper formatting of numerical, percentage, currency, and date fields.

### 📊 Pivot Table Analysis

- Total Sales by Region.
- Sales and Profit by Product Category.
- Total Quantity by Salesperson.
- Sales by Customer Segment and Payment Mode.

### 🧮 Advanced Excel Formulas

- `VLOOKUP`
- `INDEX-MATCH`
- `IF`
- `Nested IF`
- `SUM`
- `SUMIF`
- `COUNTA`

### 📈 Data Visualization

- Total Sales by Region Column Chart.
- Total Sales and Profit by Product Category Chart.
- Total Quantity by Salesperson Bar Chart.
- Conditional Formatting for High Value Orders.

### 🎛️ Interactive Analysis

- Region Slicer.
- Order Date Timeline.
- Pivot Table filtering and analysis.

### 📋 Dashboard

- Total Sales KPI.
- Total Profit KPI.
- Total Quantity KPI.
- Total Orders KPI.
- Multiple business charts.
- Clean and structured dashboard layout.

---

# 📁 Project Structure

`Excel-Sales-Performance-Dashboard/`

- `README.md`
- `SnehaGupta_ExcelDashboard.xlsx`

### Workbook Sheets

1. `Dataset` — Contains the cleaned sales dataset.
2. `Pivot Tables` — Contains Pivot Tables, Slicer and Timeline.
3. `Main Task` — Contains formulas, calculations and charts.
4. `Dashboard` — Contains KPIs and visual charts.

---

# 🔄 Project Workflow

```text
                           START
                             │
                             ▼
                  ┌──────────────────────┐
                  │   Select Dataset     │
                  └──────────┬───────────┘
                             │
                             ▼
                  ┌──────────────────────┐
                  │ Data Cleaning &      │
                  │ Preparation          │
                  └──────────┬───────────┘
                             │
                             ▼
                  ┌──────────────────────┐
                  │ Create Pivot Tables  │
                  └──────────┬───────────┘
                             │
                             ▼
                ┌──────────────────────────┐
                │ Advanced Excel Formulas │
                └────────────┬─────────────┘
                             │
               ┌─────────────┼─────────────┐
               ▼             ▼             ▼
            VLOOKUP      INDEX-MATCH     IF / Nested IF
               │             │             │
               └─────────────┼─────────────┘
                             ▼
                  ┌──────────────────────┐
                  │ Data Visualization   │
                  └──────────┬───────────┘
                             │
                             ▼
                  ┌──────────────────────┐
                  │ Interactive          │
                  │ Dashboard            │
                  └──────────┬───────────┘
                             │
                             ▼
                  ┌──────────────────────┐
                  │ Documentation        │
                  └──────────┬───────────┘
                             │
                             ▼
                        FINAL REPORT
```

---

# 🧹 Part A — Data Cleaning & Preparation

### 1. Dataset Import

The sales dataset was imported into Microsoft Excel and organized into a structured worksheet named:

`Dataset`

The dataset contains **228 sales records** across 13 columns.

### 2. Excel Table

The dataset was converted into an Excel Table named:

`SalesData`

This provides structured formatting, filtering, easier formula handling, and improved data management.

### 3. Data Formatting

Important columns were formatted according to their data types:

| Column | Format |
|--------|--------|
| `Unit_Price` | Number / Currency |
| `Quantity` | Number |
| `Discount` | Percentage |
| `Order_Date` | Date |
| `Sales` | Number / Currency |
| `Profit` | Number / Currency |

### 4. Duplicate Check

The dataset was checked for duplicate records using Excel's **Remove Duplicates** feature.

### 5. Data Validation

The dataset was reviewed to ensure that important fields were correctly formatted and suitable for Pivot Table and formula-based analysis.

---

# 📊 Part B — Pivot Table Analysis

Pivot Tables were created to summarize and analyze the sales dataset.

### 1. Total Sales by Region

**Objective:** Analyze total sales generated by each region.

`Region → Total Sales`

### 2. Sales & Profit by Product Category

**Objective:** Compare total sales and total profit for each product category.

`Product Category → Total Sales + Total Profit`

### 3. Quantity by Salesperson

**Objective:** Analyze the total quantity handled by each salesperson.

`Salesperson → Total Quantity`

### 4. Sales by Customer Segment & Payment Mode

**Objective:** Analyze sales across customer segments and payment modes.

`Customer Segment × Payment Mode → Sales`

### Interactive Filtering

A **Region Slicer** and **Order Date Timeline** were added to the Pivot Tables sheet to allow interactive filtering of Pivot Table data.

---

# 🧮 Part C — Advanced Excel Formulas

Advanced Excel formulas were used to perform lookup operations, classification, and sales categorization.

### 1. VLOOKUP

**Objective:** Use VLOOKUP to retrieve the Region Manager based on the Region.

Formula used:

`=IFERROR(VLOOKUP(B4,$E$4:$F$8,2,FALSE),"Not Found")`

The formula searches for the region and returns the corresponding Region Manager.

### 2. INDEX-MATCH

**Objective:** Use INDEX-MATCH to retrieve the Region Target based on the Region.

Formula used:

`=IFERROR(INDEX($G$4:$G$8,MATCH(B235,$E$4:$E$8,0)),"Not Found")`

`MATCH` finds the position of the region, while `INDEX` returns the corresponding target value.

### 3. IF Function

**Objective:** Classify each order as High Value or Regular based on Sales.

Formula used:

`=IF(B467>=25000,"High Value","Regular")`

Orders with Sales greater than or equal to `25000` are classified as **High Value**, while the remaining orders are classified as **Regular**.

### 4. Nested IF

**Objective:** Categorize sales into different levels.

Formula used:

`=IF(B700>=50000,"Very High",IF(B700>=25000,"High",IF(B700>=10000,"Medium","Low")))`

The sales categories used are:

- `Very High`
- `High`
- `Medium`
- `Low`

---

# 📈 Part D — Data Visualization

Charts and Conditional Formatting were used to present important sales information visually.

### 1. Total Sales by Region

A **Clustered Column Chart** was created to display Total Sales by Region.

**Analysis:**  
The chart provides a visual comparison of sales performance across the five regions.

### 2. High Value Orders

Conditional Formatting was applied to the Sales column to highlight orders with Sales greater than `25000`.

**Purpose:**  
To quickly identify High Value orders within the dataset.

### 3. Total Sales & Profit by Product Category

A **Clustered Column Chart** was created to compare Total Sales and Total Profit across product categories.

**Purpose:**  
To visually compare the financial performance of different product categories.

### 4. Total Quantity by Salesperson

A **Bar Chart** was created to show Total Quantity handled by each salesperson.

**Purpose:**  
To compare salesperson-level quantity performance.

---

# 📋 Part E — Interactive Dashboard

An interactive Sales Dashboard was created to present key business metrics and visual insights in one place.

### KPI Cards

The dashboard contains four KPI Cards:

| KPI | Purpose |
|-----|---------|
| **Total Sales** | Displays the overall sales generated |
| **Total Profit** | Displays the overall profit generated |
| **Total Quantity** | Displays the total quantity of products sold |
| **Total Orders** | Displays the total number of orders |

### Dashboard Charts

The Dashboard includes:

- **Total Sales by Region**
- **Total Sales and Profit by Product Category**
- **Total Quantity by Salesperson**

These charts provide a quick visual overview of sales performance.

### Interactive Controls

The workbook also contains:

- **Region Slicer**
- **Order Date Timeline**

These controls are available on the `Pivot Tables` sheet for interactive Pivot Table analysis.

---

# 📝 Part F — Documentation

The project documentation was created to clearly explain the workbook structure and analysis process.

### Workbook Sheets

- `Dataset` — Contains the cleaned sales dataset.
- `Pivot Tables` — Contains Pivot Tables, Slicer and Timeline.
- `Main Task` — Contains formulas, calculations and charts.
- `Dashboard` — Contains KPIs and visual charts.

### Key Excel Functions Used

`SUM`, `SUMIF`, `COUNTA`, `IF`, `Nested IF`, `VLOOKUP`, `INDEX-MATCH`

### Dashboard Components

`KPI Cards`, `Column Charts`, `Bar Chart`, `Slicer`, and `Timeline`

### Instructions

Use the **Slicer** and **Timeline** on the `Pivot Tables` sheet to filter Pivot Table data.

The Dataset contains **228 sales records** and has been cleaned and formatted for analysis.

### Naming Convention

Sheets and charts are given clear and descriptive names for easy identification.

---

# 🛠️ Tech Stack

| Tool / Technology | Usage |
|-------------------|-------|
| **Microsoft Excel** | Data analysis and dashboard development |
| **Excel Tables** | Structured data management |
| **Pivot Tables** | Data summarization and analysis |
| **Advanced Formulas** | Lookup and data classification |
| **Conditional Formatting** | Highlighting important records |
| **Charts** | Data visualization |
| **Slicer** | Interactive filtering |
| **Timeline** | Date-based filtering |

---

# 📊 Results & Learning Outcomes

Through this project, the following practical skills were developed:

- Data cleaning and preparation in Excel.
- Creating and managing structured Excel Tables.
- Working with Pivot Tables.
- Using advanced Excel lookup formulas.
- Applying logical and nested conditions.
- Creating professional data visualizations.
- Using Conditional Formatting for analysis.
- Creating KPI Cards.
- Designing an interactive dashboard.
- Using Slicers and Timelines for interactive filtering.
- Organizing and documenting an Excel data analysis project.

The project demonstrates how Excel can be used to transform structured sales data into an organized and visually understandable business report.

---

# ⭐ Advantages

- Provides a structured approach to sales data analysis.
- Makes large datasets easier to understand.
- Reduces manual analysis using Excel formulas.
- Provides quick access to important KPIs.
- Helps compare sales performance across different categories.
- Supports interactive filtering through Slicers and Timelines.
- Improves data presentation through charts and dashboards.
- Demonstrates practical business and data analytics skills.

---

# 📜 License

This project is created for **educational and portfolio purposes**.

You are free to view and learn from the project structure and analysis approach.

---

# 👩‍💻 Author

<div align="center">

### **Sneha Gupta**

**BBA Student | Data Analytics Learner | Business & Data Analytics Enthusiast**

📍 India

**Skills:**  
`Excel` · `Data Analysis` · `Power BI` · `SQL` · `Python` · `Business Analytics`

</div>

---

# 🙏 Acknowledgements

- Microsoft Excel for providing powerful tools for data analysis and visualization.
- Dataset resources used for practicing sales data analysis.
- Learning resources and practical exercises that helped develop Excel and data analytics skills.

---

<div align="center">

### ⭐ Thank You for Visiting This Project!

**Excel • Data Analysis • Visualization • Dashboard**

</div>
