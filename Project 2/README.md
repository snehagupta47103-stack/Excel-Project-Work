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

<img width="1368" height="395" alt="Screenshot 2026-09-12 131213" src="https://github.com/user-attachments/assets/266ff266-f7bf-45c5-adbb-acaf898ea3c6" />

### Summary

The Pivot Table summarizes total sales by both **Region** and **Product Category**, making it easy to compare regional performance and identify strong-performing product categories.

## `Pivot Table — Total Sales by Region`
<img width="1351" height="497" alt="Screenshot 2026-09-12 131257" src="https://github.com/user-attachments/assets/0987d8ca-743c-461c-84d5-fbccf1c6f8b7" />

## `Pivot Table - Total Sales by Product Category`
<img width="1107" height="430" alt="Screenshot 2026-09-12 131410" src="https://github.com/user-attachments/assets/52202972-4d15-40c5-a70c-a27899e7aa9a" />


---

# 9️⃣ Data Visualizations

Three major charts were created to visually communicate the sales data:

- Bar Chart
- Line Chart
- Pie Chart

These charts convert numerical data into easy-to-understand visual insights.

---

# 📊 Bar Chart — Total Sales by Region

### Summary

The Bar Chart compares the total sales generated by each region.

It helps quickly identify which regions have the highest and lowest sales performance.

The regional totals analyzed include:

- Central
- East
- North
- South
- West

## `Bar Chart — Total Sales by Region`
<img width="777" height="390" alt="Screenshot 2026-09-12 131952" src="https://github.com/user-attachments/assets/f0a91ae6-ce2b-4d14-85bb-3fd218a1b48a" />


---

# 📈 Line Chart — Monthly Sales Trend

### Summary

The Line Chart displays the movement of sales over time.

It helps identify:

- 📈 Sales increases
- 📉 Sales decreases
- 📅 Monthly trends
- 🔄 Changes in sales performance over the analyzed period

This visualization is useful for understanding the overall sales trend across different months.

## `Line Chart — Monthly Sales Trend`
<img width="1746" height="523" alt="Screenshot 2026-09-12 132108" src="https://github.com/user-attachments/assets/dee5f79b-2bce-4427-8885-b6ae8f41cc46" />


---

# 🥧 Pie Chart — Total Sales by Product Category

### Summary

The Pie Chart shows the contribution of each product category to total sales.

The categories include:

- 📚 Books
- 👕 Clothing
- 💻 Electronics
- 🪑 Furniture
- 📎 Office Supplies

This visualization makes it easy to compare the relative contribution of each product category.

## `Pie Chart — Total Sales by Product Category`
<img width="1231" height="603" alt="Screenshot 2026-09-12 132146" src="https://github.com/user-attachments/assets/d133fd1a-4bc3-4024-8a9a-1c8f90527e9a" />

---

# 📊 Dashboard

A professional **Sales & Profit Dashboard** was created to bring important KPIs and visualizations together in one place.

The dashboard includes:

### Key Performance Indicators

- **Total Sales:** 195,217.80
- **Total Profit:** 68,287.01
- **Total Quantity:** 1,999
- **Average Discount:** 9.725%

### Interactive Region Selection

A **Data Validation dropdown** was added for selecting a region.

The dashboard dynamically displays the selected region in the title, for example:

**Total Sales by Region - East**

---

# 📖 Data Storytelling

The dashboard converts raw sales transactions into a meaningful business story.

Instead of looking at hundreds of individual transactions, management can quickly understand:

- 📊 Overall sales performance
- 💰 Overall profitability
- 🌍 Regional sales contribution
- 📦 Product-category contribution
- 📈 Monthly sales movement
- 👥 Customer value
- 🏷️ Discount impact
- 🎯 Key performance indicators

This makes the analysis more useful for **business decision-making** by presenting important information in a clear, visual, and easy-to-understand format.

---

# 🔍 Key Business Insights

Based on the analysis, several useful observations can be made:

### 💰 Sales Performance

The dataset generated approximately **195K in total sales**, indicating the overall sales volume across the analyzed transactions.

### 📈 Profitability

The total profit is approximately **68K**, showing that the transactions generated a significant positive contribution.

### 🌎 Regional Performance

Regional analysis shows that sales performance differs across Central, East, North, South, and West.

The Pivot Table and Bar Chart make these differences easier to compare.

### 🛍️ Product Performance

Books, Clothing, Electronics, Furniture, and Office Supplies contribute differently to total sales.

The Pie Chart provides a visual comparison of their contribution.

### 📅 Monthly Performance

The Monthly Sales Trend and Sales Growth analysis help identify periods of increasing and decreasing sales.

### 👥 Customer Analysis

Top customer and high-value customer analysis helps identify customers who contribute significantly to revenue.

---

# 🧠 Skills Demonstrated

This project demonstrates practical knowledge of:

- 📊 Data Cleaning & Organization
- 📋 Excel Tables
- ➕ Data Aggregation
- 🎨 Conditional Formatting
- 🔄 What-If Analysis
- 📈 Statistical Analysis
- 📉 Regression Analysis
- 👥 Customer Analysis
- 📌 Pivot Tables
- 📊 Data Visualization
- 📋 Dashboard Creation
- 📖 Data Storytelling
- 💼 Business Analysis
- ⚡ Interactive Excel Reporting

---

# 📁 Workbook Structure

The Excel workbook is organized into separate sheets for better usability.

### `Project Instructions`

Contains the project requirements and task instructions.

### `Dataset`

Contains the original sales transaction data organized as an Excel Table.

### `Main_Task`

Contains calculations, task outputs, analysis tables, and supporting formulas.

### `Pivot_Table`

Contains Pivot Table analysis and supporting visualizations.

### `Dashboard`

Contains the final interactive Sales & Profit Dashboard and visual KPI analysis.

---

# 🚀 Conclusion

This project demonstrates how **Microsoft Excel** can be used as a powerful data analysis and business intelligence tool.

Starting from raw sales transactions, the project applies formulas, statistical techniques, What-If Analysis, Pivot Tables, charts, and dashboard design to generate meaningful business insights.

The final dashboard provides a concise and interactive view of:

**Sales → Profit → Customers → Regions → Products → Trends**

This project strengthened practical skills in **Excel Data Analysis, Business Intelligence, Data Visualization, and Data Storytelling**.

---

# ⭐ Project Highlights

| Area | Technique |
|------|-----------|
| Customer Analysis | `SUMIF()`, `INDEX()`, `MATCH()`, `FILTER()` |
| Profit Analysis | What-If Analysis |
| Statistical Analysis | Descriptive Statistics |
| Regression | Data Analysis ToolPak |
| Monthly Analysis | `SUMIFS()`, `EDATE()` |
| Timestamp | `NOW()` |
| Regional Analysis | Pivot Table |
| Product Analysis | Pivot Table + Pie Chart |
| Visualization | Bar, Line & Pie Charts |
| Dashboard | KPI + Data Validation + Charts |
| Data Storytelling | Business Insights |

---

# 🎯 Project Outcome

This project successfully demonstrated the practical application of **Microsoft Excel for Data Analysis and Business Intelligence**.

Through this project, I was able to:

- Analyze and organize a real-world sales dataset.
- Identify top and high-value customers.
- Perform What-If Analysis to understand discount impact on profit.
- Apply Descriptive Statistics and Regression Analysis.
- Analyze monthly sales and growth.
- Create Pivot Tables for regional and product analysis.
- Build Bar, Line, and Pie Charts for data visualization.
- Create an interactive Sales & Profit Dashboard.
- Use Data Validation for interactive reporting.
- Convert raw data into meaningful business insights.
- Present analytical findings through effective Data Storytelling.

Overall, the project improved my practical skills in **Excel Data Analysis, Statistical Analysis, Data Visualization, Dashboard Development, Business Intelligence, and Data Storytelling**.

> **The final outcome is a complete and interactive Excel-based sales analysis solution that transforms raw transaction data into meaningful, decision-ready business information.**

---

---

<div align="center">

# 👩‍💻 Author

**Name:** Sneha Gupta  
**Project:** Excel Data Analysis — Project 2  
**Course:** Data Analysis  
**Tool Used:** Microsoft Excel  

This project was created as part of a practical **Excel Data Analysis** project to demonstrate skills in data analysis, statistical analysis, visualization, dashboard creation, and business storytelling.

</div>

---

