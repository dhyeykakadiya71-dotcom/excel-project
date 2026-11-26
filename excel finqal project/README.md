# 📊 Customer Transactions Analysis & Dashboard

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Excel%20Formulas-blue?style=flat)
![GitHub](https://img.shields.io/badge/GitHub-Repository-181717?style=flat&logo=github&logoColor=white)

An Excel project that analyzes 250 customer transactions to understand revenue, customer spending, product performance, and regional sales. The workbook uses Excel formulas, charts, and a summary dashboard.

**Main file:** `Final_Project.xlsx`

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Worksheet Details](#-worksheet-details)
- [Sample Project Results](#-sample-project-results)
- [Excel Functions Used](#-excel-functions-used)
- [Technologies and Tools](#-technologies-and-tools)
- [Project Structure](#-project-structure)
- [How to Use](#-how-to-use)
- [Learning Outcomes](#-learning-outcomes)
- [Future Improvements](#-future-improvements)
- [Author](#-author)

---

## 📌 Project Overview

The workbook contains **250 customer transactions** from **11 April 2024 to 11 April 2025**, covering **50 customers**, **10 products** in 3 categories (Appliances, Electronics, Furniture), and **5 regions** (Central, East, North, South, West).

It is organized into four worksheets:

| # | Worksheet | Purpose |
|---|---|---|
| 1 | **Raw Data** | Original transaction dataset plus formula-based helper columns |
| 2 | **Analysis** | Key metrics, customer and product analysis, filtering, regression, and What-If Analysis |
| 3 | **Visualizations** | Summary tables and three charts (region, month, category) |
| 4 | **Dashboard** | One-page summary with KPIs, ranked tables, a chart, and a revenue target check |

---

## 📂 Worksheet Details

### 1️⃣ Raw Data

This sheet holds the original transaction dataset (rows 5–254, stored as an Excel table named `Transactions`).

**Original columns:**

`Transaction_ID`, `Date`, `Customer_ID`, `Customer_Name`, `Product_ID`, `Product_Name`, `Category`, `Quantity`, `Unit_Price`, `Payment_Method`, `Region`, `Customer_Segment`, `Customer_Since`, `Total_Amount`

**Formula-based columns:**

| Column | What it calculates |
|---|---|
| `Create_Timestamp` | Current date and time, using `NOW()` |
| `Days_Since_Transaction` | Days between today and the transaction date |
| `Customer_Tenure_Years` | Completed years since `Customer_Since` |
| `Month_End` | Last day of the transaction's month |
| `Customer_Initials` | First letter of the first and last name, in capitals (e.g. Paul Baker → PB) |
| `Region_Category_Sequence` | A running key such as `Central\|Electronics\|1`, which numbers each transaction within its region and category |

> `Create_Timestamp`, `Days_Since_Transaction`, and `Customer_Tenure_Years` depend on today's date, so they change whenever the workbook is recalculated. The sheet also uses a data bar on `Days_Since_Transaction`.

### 2️⃣ Analysis

**Key Performance Metrics**

Total Revenue, Total Transactions, Total Quantity Sold, Unique Customers, Average Transaction Value, First and Last Transaction Date, Most Frequently Purchased Product, and Most Frequent Customer.

**Customer Analysis**

- A customer-level table for all 50 customers showing **total spend** and **number of transactions**.
- Each customer is labeled **High Value** if their total spend is at or above the average customer spend, otherwise **Regular**. In the current data this gives 24 High Value and 26 Regular customers.
- A separate list of transaction counts per customer name.
- Data bars highlight total spend and icons highlight transaction counts.

**Product Analysis**

- Transaction counts for each of the 10 products.
- The most frequently purchased product is found automatically from these counts.

**Multi-value Filtering**

- A results table that lists every matching transaction (ID, customer, product, revenue) for a **region** and **category** chosen from two dropdown lists.
- It is built with `INDEX`/`MATCH` and the `Region_Category_Sequence` helper column from Raw Data. The section is labelled "FILTER-style", but it does not use Excel's `FILTER` function.

**Linear Regression (Quantity vs Total Amount)**

- Slope, intercept, and R-squared using `SLOPE`, `INTERCEPT`, and `RSQ`.
- An auto-generated regression equation.
- A short written interpretation: quantity alone does not fully explain revenue because products have different unit prices.

**What-If Analysis**

- Average order value assumption (linked to the average transaction value)
- Target revenue (250,000)
- Required transaction count (target revenue ÷ average order value)
- Written instructions for using **Goal Seek** and **Scenario Manager** (Base, +10% Price, −10% Price). These are instructions in the sheet; no Goal Seek or Scenario Manager results are saved in the workbook.

### 3️⃣ Visualizations

This sheet holds the supporting data tables and the three charts built from them:

| Supporting data | Chart |
|---|---|
| Revenue by Region | Bar chart — *Revenue by Region* |
| Revenue by Month (Apr 2024 – Apr 2025) | Line chart — *Monthly Revenue Trend* |
| Revenue by Category | Pie chart — *Revenue by Category* |

The charts make it easier to compare regions, spot ups and downs across months, and see how revenue splits between categories without reading the tables row by row.

### 4️⃣ Dashboard

The main summary sheet.

**Key Performance Indicators:** Total Revenue, Transactions, Customers, and Units Sold.

**Dashboard analysis:**

- **Top Customers** — the top 5 customers by total spending, ranked automatically.
- **Regional Sales** — revenue for each of the 5 regions, with a bar chart (*Regional Revenue*) beside it.
- **Most Sold Products** — units sold for each of the 10 products.
- **Revenue Target Check** — target revenue, current revenue, and a status that shows either "Target reached" or "Below target".

Data bars and icon sets are used on the dashboard tables to make high and low values easier to spot.

---

## 📈 Sample Project Results

These are the results for the **current sample dataset** in the workbook. They are not guaranteed results for any future or different data.

| Metric | Result |
|---|---:|
| Total Revenue | 229,192.47 |
| Total Transactions | 250 |
| Unique Customers | 50 |
| Total Quantity Sold | 753 |
| Average Transaction Value | 916.77 |
| Revenue Target | 250,000 |

With the current data, revenue is below the 250,000 target, and the dashboard status shows **Below target**.

---

## 🧮 Excel Functions Used

### Functions requested for this README

| Function | What it does | How it is used in this project |
|---|---|---|
| `NOW()` | Returns the current date and time | `Create_Timestamp` column in Raw Data |
| `TODAY()` | Returns today's date | Calculates `Days_Since_Transaction` and customer tenure |
| `IF()` | Returns different results based on a condition | Labels customers High Value or Regular; adjusts tenure if the anniversary has not occurred yet; sets the target status on the Dashboard |
| `IFERROR()` | Replaces an error with a fallback value | Used in tenure, customer name lookups, and the multi-value filtering table so errors and empty rows stay clean |
| `YEAR()` | Extracts the year from a date | Part of the customer tenure calculation |
| `DATE()` | Builds a date from year, month, and day | Builds this year's anniversary date for the tenure calculation |
| `MONTH()` | Extracts the month from a date | Part of the customer tenure calculation |
| `DAY()` | Extracts the day from a date | Part of the customer tenure calculation |
| `EOMONTH()` | Returns the last day of a month | `Month_End` column |
| `UPPER()` | Converts text to capital letters | `Customer_Initials` column |
| `LEFT()` | Returns characters from the start of text | Takes the first letter of the customer's name |
| `MID()` | Returns characters from the middle of text | Takes the first letter of the last name |
| `FIND()` | Finds the position of a character in text | Finds the space between first and last name |
| `COUNTIFS()` | Counts rows matching multiple conditions | Creates the running number in `Region_Category_Sequence` |

### Other functions also used in the workbook

| Function | Where it is used |
|---|---|
| `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNTA` | Key metrics and dashboard KPIs |
| `COUNTIF` | Transaction counts per product, per customer name, and per customer ID |
| `SUMIF`, `SUMIFS` | Revenue by customer, region, category, and month; units sold per product |
| `INDEX`, `MATCH` | Name lookups, most frequent product/customer, multi-value filtering |
| `LARGE`, `ROWS` | Ranking top customers on the Dashboard |
| `EDATE` | Month boundaries in the monthly revenue calculation |
| `SLOPE`, `INTERCEPT`, `RSQ`, `TEXT` | Linear regression and the regression equation text |

### Analysis techniques

- **What-If Analysis** with a target revenue and required transaction count.
- **Goal Seek** and **Scenario Manager**, documented as instructions in the Analysis sheet.

---

## 🛠️ Technologies and Tools

- Microsoft Excel
- Excel formulas and functions
- Data analysis
- Charts and visualizations
- Dashboard creation
- Conditional formatting (data bars and icon sets)
- Data validation (dropdown lists for region and category)

---

## 📁 Project Structure

```
Excel-Project/
│
├── Final_Project.xlsx
└── README.md
```

---

## ▶️ How to Use

1. Clone or download the repository:
   ```bash
   git clone https://github.com/dhyeykakadiya71-dotcom/excel-project.git
   ```
2. Open the project folder.
3. Open `Final_Project.xlsx` in Microsoft Excel.
4. Move between the **Raw Data**, **Analysis**, **Visualizations**, and **Dashboard** sheets.
5. Explore the transaction data, analysis results, charts, and dashboard.

**Tip:** On the Analysis sheet, change the Region (`H17`) and Category (`H18`) dropdowns to see the matching transactions update.

---

## 🎯 Learning Outcomes

This project demonstrates:

- Working with structured datasets
- Excel formulas and functions
- Customer transaction analysis
- Revenue calculations
- Customer segmentation (High Value vs Regular)
- Product performance analysis
- Regional sales analysis
- Date and time calculations
- Linear regression basics
- What-If Analysis
- Chart creation
- Dashboard development
- Data interpretation and decision-making

---

## 🚀 Future Improvements

These are ideas for later and are **not** part of the current workbook:

- Interactive dashboard filters
- Automatic data updates
- More advanced customer segmentation
- Additional sales performance metrics
- Improved forecasting
- Expanded visualizations

---

## 👤 Author

**Dhyey Kakadiya**

🔗 [GitHub Repository](https://github.com/dhyeykakadiya71-dotcom/excel-project)
