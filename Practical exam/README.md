# 📊 Excel Practical Exam – Customer Sales Analysis

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Sales%20%26%20Profit-blue?style=flat)
![GitHub](https://img.shields.io/badge/GitHub-Repository-181717?style=flat&logo=github&logoColor=white)

**Author:** Dhyey Kakadiya  
**Excel File:** `Practical_Exam.xlsx`  
**Software:** Microsoft Excel

---

## 📌 Project Overview

This project is an Excel-based **Customer Sales and Profit Analysis**, created as part of a practical examination.

It shows how Microsoft Excel can be used to organize customer transaction data, clean and process it, retrieve customer details with lookup formulas, summarize sales performance, and present the results through charts.

The workbook contains **200 customer transaction records** spread across **five worksheets**: `Raw_Data`, `Clean_Data`, `Pivot_Table`, `Advanced_Formulas` and `Dashboard`. The order dates in the data fall in 2024 and 2025.

---

## 🗂️ Worksheet Details

### 1️⃣ Raw_Data

The original customer transaction dataset. It is the source for all the other sheets and has these columns:

- `Customer_ID`
- `Customer_Name`
- `Region`
- `Product_Category`
- `Sales`
- `Quantity`
- `Discount`
- `Order_Date`
- `Profit`

### 2️⃣ Clean_Data

The processed version of the dataset. It keeps the original columns and adds four calculated columns:

- **Month** – extracted from `Order_Date` (for example, *Jan*, *Feb*) using `TEXT()`
- **Year** – extracted from `Order_Date` using `TEXT()`
- **Sales_Status** – classifies each transaction using `IF()`
- **Profit_Status** – classifies each transaction using a nested `IF()`

**Classification logic**

| Column | Category | Rule |
|---|---|---|
| Sales_Status | High Sales | Sales ≥ 1000 |
| Sales_Status | Regular Sales | Sales < 1000 |
| Profit_Status | High | Profit ≥ 400 |
| Profit_Status | Medium | Profit ≥ 200 and < 400 |
| Profit_Status | Low | Profit < 200 |

### 3️⃣ Pivot_Table

This sheet holds **three PivotTables**, all built from the `Clean_Data` table.

**Regional Analysis**
- Region
- Sum of Sales
- Sum of Profit
- Grand Total

**Product Category Analysis**
- Product Category
- Sum of Sales
- Sum of Quantity
- Grand Total

**Monthly Analysis**
- Month
- Sum of Sales
- Sum of Profit
- Grand Total

Together, these summaries make it easy to compare regional performance, see which product categories sell the most, and follow how sales and profit change from month to month. The monthly summary groups by month name, so the same month from 2024 and 2025 is added together.

### 4️⃣ Advanced_Formulas

A customer lookup section that demonstrates lookup and error-handling formulas. It covers the first 19 customers (`CUST001` to `CUST019`) and has four columns:

| Column | What it does |
|---|---|
| Search Customer_ID | Pulls the customer ID directly from `Raw_Data` |
| Customer_Name | Finds the customer's name with `VLOOKUP()` |
| Customer_Region | Finds the customer's region with `INDEX()` and `MATCH()` |
| Customer_Lookup_Check | Repeats the name lookup inside `IFERROR()`, so a missing ID shows "Customer Not Found" instead of an error |

| Excel Function | Purpose |
|---|---|
| VLOOKUP | Retrieves customer information |
| INDEX | Returns a value from a selected range |
| MATCH | Finds the position of a matching value |
| IFERROR | Handles errors in formula results |

### 5️⃣ Dashboard

The visual summary of the project. It contains **three charts** and **three slicers**:

1. **Regional Sales and Profit Analysis** – a column chart comparing sales and profit across regions.
2. **Product Category Sales Analysis** – a bar chart showing sales (alongside quantity) for each product category.
3. **Monthly Sales and Profit Trend** – a line chart showing monthly sales and profit.

The charts are linked to the PivotTables on the `Pivot_Table` sheet. The dashboard also has slicers for **Region**, **Product_Category** and **Month**. Each slicer filters the PivotTable and chart it is connected to, which makes it quick to focus on a single region, category or month.

---

## 📈 Project Results

These figures come from the current workbook (`Practical_Exam.xlsx`):

| Metric | Result |
|---|---:|
| Total Transactions | 200 |
| Total Sales | 192,748.84 |
| Total Profit | 67,341.45 |

---

## 🧮 Excel Functions Used

| Function | What it does | How it is used in this project |
|---|---|---|
| `TEXT()` | Converts a value into text in a chosen format | Extracts the month name and the year from `Order_Date` in `Clean_Data` |
| `IF()` | Returns one result if a condition is true and another if it is false | Sets `Sales_Status` to *High Sales* or *Regular Sales* |
| Nested `IF()` | Places one `IF()` inside another to handle more than two outcomes | Sets `Profit_Status` to *High*, *Medium* or *Low* |
| `VLOOKUP()` | Looks up a value in the first column of a table and returns a value from another column | Finds the customer name for each customer ID |
| `INDEX()` | Returns the value at a given position in a range | Returns the customer's region from the `Region` column |
| `MATCH()` | Finds the position of a value in a range | Finds the row of the customer ID so `INDEX()` can return the region |
| `IFERROR()` | Replaces a formula error with a custom result | Shows "Customer Not Found" if a lookup fails |

---

## 🛠️ Technologies Used

- Microsoft Excel
- Excel Formulas and Functions
- Data Cleaning
- Data Analysis
- Summary Tables (PivotTables)
- Data Visualization
- Excel Charts

---

## 📁 Project Structure

```
Practical exam/
│
├── Practical_Exam.xlsx
└── README.md
```

---

## ▶️ How to Use the Project

1. Open the GitHub repository.
2. Download or clone the repository.
3. Open the `Practical_Exam.xlsx` file in Microsoft Excel.
4. Explore all five worksheets.
5. Review the original data, cleaned data, summaries, advanced formulas, and dashboard.

> 💡 The slicers on the Dashboard need Excel 2010 or later to work.

---

## 🎓 Learning Outcomes

This project demonstrates:

- Organizing customer transaction datasets
- Cleaning and processing spreadsheet data
- Using conditional logic
- Extracting months and years from dates
- Categorizing sales and profit
- Creating summarized reports with PivotTables
- Understanding `VLOOKUP`
- Using `INDEX` and `MATCH` together
- Handling formula errors
- Analyzing regional and product performance
- Understanding monthly sales trends
- Creating charts, slicers and a dashboard

---

## 🚀 Future Improvements

These are ideas for the future, not features of the current workbook:

- Adding more dashboard filters, such as a Year slicer
- Expanding the customer dataset
- Extending the lookup section beyond the first 19 customers
- Adding more sales performance metrics
- Including additional charts
- Automating data updates
- Introducing more advanced Excel analysis

---

## 👤 Author

**Created by:** Dhyey Kakadiya

**GitHub Repository:** [excel-project / Practical exam](http://github.com/dhyeykakadiya71-dotcom/excel-project/tree/main/Practical%20exam)
