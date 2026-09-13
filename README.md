# Excel-data-exploration-assignment-1
A beginner-level data analytics project using Microsoft Excel to explore product data, perform calculations, categorize records, apply conditional functions, and extract structured information from Product IDs.
###  Project Overview

This project is part of the **Program in AI-Driven Data Analytics** course offered by **Entri App**.

The objective of this assignment is to develop fundamental Excel data analysis skills by exploring a product dataset and applying commonly used Excel functions for:

- Basic numerical analysis
- Statistical calculations
- Conditional categorization
- Conditional aggregation
- Text extraction and formatting

The assignment demonstrates how Excel functions can be used to transform raw data into meaningful information and create additional analytical columns from existing data.

---

##  Workbook Structure

The Excel workbook contains two sheets:

### 1. Instructions

This sheet contains the questions and tasks that need to be completed using the dataset.

### 2. Dataset Architecture

The dataset consists of 34 individual product records categorized across retail segments (Electronics, Fashion, Kitchen, Outdoor).

Data Schema
Product ID: Unique composite identifier containing date and country metadata (e.g., 28-JAN-US).

Product Name: Name of the retail product item.

Brand Name: Brand manufacturer.

Price ($): Unit pricing of the product.

Quantity: Total inventory/sales volume.

Category: Market sector classification.


#  Objectives

The main objectives of this assignment are to:

1. Calculate the total, count, and average of product prices.
2. Identify the minimum and maximum product prices.
3. Categorize products based on their price using the `IF` function.
4. Perform conditional calculations using `SUMIF` and `COUNTIF`.
5. Extract specific portions of Product IDs using `LEFT`, `RIGHT`, and `MID`.
6. Understand how Excel formulas can be applied to real-world data analysis tasks.

---

#  Analysis and Methodology

## 1. Sum, Count, and Average

The first step was to perform basic statistical analysis on the `Price ($)` column.

### Total Price of All Products

The `SUM` function was used to calculate the total price of all products.

**Formula:**

```excel
=SUM(D2:D35)
````

**Result:**

**$10,100**

---

### Number of Products

The `COUNT` function was used to determine the number of products in the dataset.

**Formula:**

```excel
=COUNT(D2:D35)
```

**Result:**

**34 products**

---

### Average Product Price

The `AVERAGE` function was used to calculate the average price of all products.

**Formula:**

```excel
=AVERAGE(D2:D35)
```

**Result:**

**$297.06 approximately**

---

# 2. Minimum and Maximum Price

The `MIN` and `MAX` functions were used to identify the lowest and highest product prices.

### Minimum Price

**Formula:**

```excel
=MIN(D2:D35)
```

**Result:**

**$30**

### Maximum Price

**Formula:**

```excel
=MAX(D2:D35)
```

**Result:**

**$1,000**

---

# 3. Price Range Using IF Function

A new column named **Price Range** was created to classify products based on their price.

According to the assignment:

* Products with a price **greater than or equal to $500** → `High Price`
* Products with a price **below $500** → `Standard Price`

### Formula

```excel
=IF(D2>=500,"High Price","Standard Price")
```

The formula was entered in the first row of the `Price Range` column and filled down for all products.

### Logic

The formula checks the product price in column D.

If the price is at least $500, Excel returns:

```text
High Price
```

Otherwise, it returns:

```text
Standard Price
```

This demonstrates the use of a logical condition to create a new categorical variable from numerical data.

---

# 4. SUMIF and COUNTIF

Conditional functions were used to perform more specific analysis on the dataset.

## Electronics Category – Total Price

The `SUMIF` function was used to calculate the total price of products belonging to the **Electronics** category.

**Formula:**

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

**Result:**

**$8,050**

### Explanation

* `F2:F35` → Category range
* `"Electronics"` → Condition
* `D2:D35` → Values to be added

The formula adds the prices only for products where the category is Electronics.

---

## Products Priced Below $100

The `COUNTIF` function was used to count products with a price below $100.

**Formula:**

```excel
=COUNTIF(D2:D35,"<100")
```

**Result:**

**11 products**

### Explanation

The formula checks the `Price ($)` column and counts every product whose price is less than $100.

---

# 5. Text Extraction Using LEFT, RIGHT, and MID

The `Product ID` column contains structured information in the following format:

```text
28-JAN-US
```

The Product ID consists of:

* First 2 characters → Day
* Characters 4–6 → Month
* Last 2 characters → Country Code

Excel text functions were used to extract these individual components.

---

## Day – LEFT Function

A new column named **Day** was created by extracting the first two characters from the Product ID.

**Formula:**

```excel
=LEFT(A2,2)
```

### Example

```text
Product ID: 28-JAN-US
Result: 28
```

The `LEFT` function extracts characters starting from the left side of the text.

---

## Country Code – RIGHT Function

A new column named **Country Code** was created by extracting the final two characters from the Product ID.

**Formula:**

```excel
=RIGHT(A2,2)
```

### Example

```text
Product ID: 28-JAN-US
Result: US
```

The `RIGHT` function extracts characters from the right side of the text.

---

## Month – MID Function

A new column named **Month** was created by extracting characters 4 to 6 from the Product ID.

**Formula:**

```excel
=MID(A2,4,3)
```

### Example

```text
Product ID: 28-JAN-US
Result: JAN
```

The `MID` function extracts a specified number of characters starting from a specified position.

---

# 📊 Final Analysis Results

| Analysis                | Excel Function |  Result |
| ----------------------- | -------------- | ------: |
| Total Price             | `SUM`          | $10,100 |
| Number of Products      | `COUNT`        |      34 |
| Average Price           | `AVERAGE`      | $297.06 |
| Minimum Price           | `MIN`          |     $30 |
| Maximum Price           | `MAX`          |  $1,000 |
| Electronics Total Price | `SUMIF`        |  $8,050 |
| Products Below $100     | `COUNTIF`      |      11 |

---

# 🧮 Excel Functions Used

The following Excel functions were used to complete the assignment:

| Function  | Purpose                                     | Formula                                      |
| --------- | ------------------------------------------- | -------------------------------------------- |
| `SUM`     | Calculates the total of product prices      | `=SUM(D2:D35)`                               |
| `COUNT`   | Counts the number of numeric product prices | `=COUNT(D2:D35)`                             |
| `AVERAGE` | Calculates the average product price        | `=AVERAGE(D2:D35)`                           |
| `MIN`     | Finds the lowest product price              | `=MIN(D2:D35)`                               |
| `MAX`     | Finds the highest product price             | `=MAX(D2:D35)`                               |
| `IF`      | Categorizes products according to price     | `=IF(D2>=500,"High Price","Standard Price")` |
| `SUMIF`   | Calculates the total price for Electronics  | `=SUMIF(F2:F35,"Electronics",D2:D35)`        |
| `COUNTIF` | Counts products priced below $100           | `=COUNTIF(D2:D35,"<100")`                    |
| `LEFT`    | Extracts the first two characters           | `=LEFT(A2,2)`                                |
| `RIGHT`   | Extracts the last two characters            | `=RIGHT(A2,2)`                               |
| `MID`     | Extracts the month from the Product ID      | `=MID(A2,4,3)`                               |

---

# 💡 Key Learnings

Through this assignment, I gained practical experience in using Excel for basic data exploration and transformation.

The key concepts covered were:

* Working with structured datasets in Excel
* Performing basic numerical calculations
* Using statistical functions to summarize data
* Applying logical conditions with the `IF` function
* Performing conditional aggregation using `SUMIF`
* Counting records based on conditions using `COUNTIF`
* Extracting meaningful information from text using `LEFT`, `RIGHT`, and `MID`
* Creating new analytical columns from existing data
* Interpreting numerical results to understand a dataset

---

#  Conclusion

This assignment demonstrates how Excel can be used as a basic data analysis tool to explore, summarize, categorize, and transform raw data.

By applying functions such as `SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX`, `IF`, `SUMIF`, `COUNTIF`, `LEFT`, `RIGHT`, and `MID`, the dataset was converted into a more informative and analysis-ready format.

These Excel fundamentals provide a strong foundation for progressing toward more advanced data analytics concepts, including data cleaning, data visualization, PivotTables, dashboards, SQL, Python, and other analytical tools.

---

##  Project Files

* `Excel Assignment 1 - Data Exploration.xlsx` – Completed Excel assignment
* `README.md` – Project overview, methodology, formulas, results, and learnings

---

## 🛠️ Tools Used

* **Microsoft Excel**
* **GitHub**

