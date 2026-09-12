# 📊 Sales & Profit Analyzer — Excel Data Analysis Project 2

> A comprehensive Microsoft Excel data analysis project focused on sales, profit, customer behavior, regional performance, product-category analysis, What-If Analysis, statistical analysis, Pivot Tables, charts, and an interactive dashboard.

---

## 📌 Project Overview

This project is designed to analyze a sales dataset using advanced Microsoft Excel features and data analysis techniques.

The project transforms raw transactional data into meaningful business insights through:

- Conditional Formatting
- What-If Analysis
- Data Analysis ToolPak
- Linear Regression
- Descriptive Statistics
- Monthly Sales Growth Analysis
- Timestamp Creation
- Customer Value Analysis
- Pivot Tables
- Bar Charts
- Line Charts
- Pie Charts
- Interactive Dashboard
- Data Storytelling

The final outcome is an Excel-based **Sales & Profit Dashboard** that provides a clear and visual overview of business performance.

---

## 🎯 Project Objectives

The main objectives of this project are:

1. Identify the top customers based on total purchase.
2. Analyze the impact of discount changes on total profit.
3. Perform Linear Regression between Sales and Profit.
4. Generate Descriptive Statistics for numerical variables.
5. Analyze monthly sales growth using up/down indicators.
6. Create timestamps using the `NOW()` function.
7. Identify high-value customers using lookup and filtering techniques.
8. Analyze total sales by Region and Product Category using Pivot Tables.
9. Create visualizations using Bar, Line, and Pie Charts.
10. Build a professional dashboard to communicate important business insights.

---

# 📂 Dataset Information

The dataset contains **200 sales transactions**.

### Dataset Columns

| Column | Description |
|---|---|
| Customer_ID | Unique customer identifier |
| Customer_Name | Customer name |
| Region | Sales region |
| Product_Category | Category of the purchased product |
| Sales | Sales amount |
| Quantity | Quantity purchased |
| Discount | Discount applied to the transaction |
| Order_Date | Date of the order |
| Profit | Profit generated from the transaction |

### Dataset Categories

**Regions:**
- Central
- East
- North
- South
- West

**Product Categories:**
- Books
- Clothing
- Electronics
- Furniture
- Office Supplies

---

# 🛠️ Tools & Excel Features Used

### Microsoft Excel Features

- Excel Tables
- Structured Data
- Conditional Formatting
- Data Validation
- What-If Analysis
- Data Analysis ToolPak
- Pivot Tables
- Charts
- Dashboard Design

### Excel Functions Used

- `SUMIF()`
- `SUMIFS()`
- `SUMPRODUCT()`
- `AVERAGE()`
- `MAX()`
- `FILTER()`
- `INDEX()`
- `MATCH()`
- `IFERROR()`
- `NOW()`
- `DATE()`
- `EDATE()`
- `UNIQUE()`
- `SORT()`

---

# 📋 Project Tasks

## 1️⃣ Top 10 Customers

Conditional Formatting was applied to identify the **Top 10 customers based on total purchase**.

The customer-level purchase values were calculated by aggregating sales for each Customer ID.

### Technique Used
- `SUMIF()`
- Conditional Formatting
- Top 10 Items rule

### Purpose

This helps identify customers who contribute the highest sales value and can therefore be considered important customers for the business.

---

# 2️⃣ What-If Analysis — Discount vs Total Profit

What-If Analysis was performed to understand how changing the discount percentage can affect total profit.

Different discount scenarios were tested:

- 0%
- 5%
- 10%
- 15%
- 20%

A Data Table was created to compare the resulting total profit under each discount scenario.

### Technique Used

- What-If Analysis
- Data Table
- `SUMPRODUCT()`

### Business Purpose

This analysis helps understand the relationship between discount decisions and profitability.

It can help management evaluate whether offering higher discounts is beneficial or whether it may reduce overall profit.

---

# 3️⃣ Linear Regression — Profit vs Sales

Linear Regression was planned using the Excel **Data Analysis ToolPak**.

### Variables

**Dependent Variable (Y):**
- Profit

**Independent Variable (X):**
- Sales

### Purpose

Regression analysis helps examine whether there is a relationship between Sales and Profit.

A positive relationship would indicate that higher sales are generally associated with higher profit.

### Excel Feature Used

**Data → Data Analysis → Regression**

---

# 4️⃣ Descriptive Statistics

Descriptive Statistics were generated using the **Analysis ToolPak**.

The numerical variables analyzed include:

- Sales
- Quantity
- Discount
- Profit

### Statistics Generated

The analysis provides statistical measures such as:

- Mean
- Standard Error
- Median
- Mode
- Standard Deviation
- Sample Variance
- Kurtosis
- Skewness
- Range
- Minimum
- Maximum
- Sum
- Count

### Purpose

Descriptive Statistics provide a quick numerical summary of the dataset and help understand the distribution and variation of the business data.

---

# 5️⃣ Monthly Sales Growth Analysis

Monthly sales were calculated from the transaction dates and compared month-to-month.

The analysis includes:

- Month
- Total Sales
- Sales Growth

Conditional Formatting with **up/down arrows** was applied to the Sales Growth column.

### Techniques Used

- `DATE()`
- `EDATE()`
- `SUMIFS()`
- Percentage Growth Formula
- Conditional Formatting
- Icon Sets

### Purpose

This makes it easier to identify months with increasing or decreasing sales performance.

---

# 6️⃣ Timestamp Creation

A timestamp column was created using:

```excel
=NOW()
```
The timestamp records the current date and time when the workbook is updated.

### Purpose

Timestamps can be useful for tracking when an analysis or report was generated.

---

# 7️⃣ High-Value Customer Analysis

High-value customers were identified using customer purchase information and filtering techniques.

### Techniques Used

- `INDEX()`
- `MATCH()`
- `FILTER()`
- `IFERROR()`

The analysis helps identify customers with comparatively high purchase values.

### Business Purpose

High-value customers can be important for:

- Customer retention
- Personalized offers
- Loyalty programs
- Targeted marketing
- Revenue growth

---

# 8️⃣ Pivot Table Analysis

A Pivot Table was created to analyze **Total Sales by Region and Product Category**.

The Pivot Table contains:

- Regions as Row Labels
- Product Categories as Column Labels
- Sum of Sales as Values
- Grand Total

### Product Categories Analyzed

- Books
- Clothing
- Electronics
- Furniture
- Office Supplies

### Regions Analyzed

- Central
- East
- North
- South
- West

### Purpose

The Pivot Table provides a quick comparison of sales performance across different regions and product categories.

---

## 📸 Pivot Table Screenshot

### Summary

The Pivot Table summarizes total sales by both **Region** and **Product Category**, making it easy to compare regional performance and identify strong-performing product categories.

### 📷 Add Screenshot Here

> **Insert the screenshot of the completed Pivot Table here.**

**Suggested screenshot:**  
`Pivot Table — Total Sales by Region and Product Category`

---

# 9️⃣ Data Visualizations

Three major charts were created to visually communicate the sales data:

- Bar Chart
- Line Chart
- Pie Chart

These charts convert numerical data into easy-to-understand visual insights.
