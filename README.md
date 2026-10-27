[README (1).md](https://github.com/user-attachments/files/34021437/README.1.md)
# Data-cleaning-and-transformation# Excel Data Cleaning & Transformation

A hands-on Excel project that takes a small, messy product-inventory dataset and cleans it into an analysis-ready table using built-in Excel features and formulas.

## Table of Contents

- [Overview](#overview)
- [Repository Contents](#repository-contents)
- [Dataset](#dataset)
- [Tasks & Approach](#tasks--approach)
- [Results](#results)
- [Skills Demonstrated](#skills-demonstrated)
- [How to Use](#how-to-use)
- [Known Limitations](#known-limitations)

## Overview

Real-world data is rarely clean. This project works through the most common data-preparation problems: missing values, inconsistent text, typos, duplicate rows, columns that need splitting or merging, and inconsistent number formats. It finishes with conditional formatting to make the cleaned data easier to read.

## Repository Contents

| File | Description |
|------|-------------|
| `Excel_Assignment_2_-_Data_Cleaning_and_Transformation.xlsx` | Workbook with three sheets (see below) |
| `README.md` | Project documentation |

**Workbook sheets**

| Sheet | Purpose |
|-------|---------|
| `Instructions` | The assignment brief (six tasks) |
| `Dataset` | The raw, uncleaned data |
| `Answer` | The cleaned data, helper columns, and written explanations |

## Dataset

The raw dataset contains **34 product records** with 6 columns:

| Column | Description | Example |
|--------|-------------|---------|
| Product ID | Encoded as `DD-MON-CC` (day, month, country code) | `28-JAN-US` |
| Product Name | Name of the product (inconsistent casing) | `laptop`, `Sneakers` |
| Brand Name | Manufacturer | `Dell` |
| Price ($) | Unit price (some missing) | `1000` |
| Quantity | Units in stock | `30` |
| Category | Product category (some missing or misspelled) | `Electronics` |

### Issues found in the raw data

- **Missing prices**: 3 products have no price (Sony Headphones, Coleman Camping Tent, Ray-Ban Sunglasses).
- **Missing categories**: 4 products have a blank category (Backpack, Sneakers, Coffee Maker, Fitness Tracker).
- **Inconsistent text casing** in Product Name (`laptop`, `smartphone`, `headphones` vs. `Sneakers`, `Coffee Maker`).
- **Typos in Category**: `Electroni` instead of `Electronics`.
- **Duplicate rows**: 3 exact duplicates (Samsonite Laptop Bag, Bose Headphones, HP Laptop).
- **Composite Product ID** that mixes a date and a country code in one field.

## Tasks & Approach

### 1. Handling Missing Values
- Counted blanks in the Price column with `=COUNTBLANK(D2:D35)` and flagged them with `=ISBLANK()`.
- **Price**: Because price varies by brand, product, and category, filling in an average would be misleading. Missing prices are labelled `"UNKNOWN"` using `=IF(ISBLANK(D2),"UNKNOWN",D2)`. Note that text values can interfere with numeric calculations, so this is best used for display or paired with a filter.
- **Category**: Missing categories are filled with `"UNKNOWN"` using `=IF(ISBLANK(F2),"UNKNOWN",F2)`. A stronger strategy would be to infer the category from similar products (e.g. a Coffee Maker is most likely Kitchen).

### 2. Correcting Inconsistent Data
- Standardised **Product Name** to proper case (e.g. `laptop` becomes `Laptop`).
- Fixed **Category** typos (`Electroni` becomes `Electronics`) using Find & Replace and the Spelling checker (Review > Spelling).

### 3. Removing Duplicates
- Used **Data > Remove Duplicates** across all columns, reducing the dataset from 34 rows to 31 unique rows.

### 4. Splitting and Merging Data
- Split **Product ID** into:
  - `Manufacturing Date`: `=DATEVALUE(LEFT(A2,6)&"-2026")`
  - `Country Code`: `=RIGHT(A2,2)`
- Merged **Product Name** and **Brand Name** into **Product Brand**: `=B2&" "&C2` (e.g. `Laptop Dell`).

### 5. Number Formatting
- Applied currency format to the **Price** column.
- Applied a day-month-year date format to the **Manufacturing Date** column.

### 6. Conditional Formatting
- Added **data bars** to the Price column to visualise relative price.
- Added a **custom rule** on the Category column to highlight cells equal to `"Electronics"`.

## Results

| Metric | Before | After |
|--------|--------|-------|
| Rows | 34 | 31 |
| Duplicate rows | 3 | 0 |
| Category typos | Multiple (`Electroni`) | 0 |
| Product Name formats | Mixed case | Consistent proper case |
| Missing prices / categories | Blank cells | Explicitly flagged as `UNKNOWN` |
| Date and country data | Combined in one column | Two separate, usable columns |

## Skills Demonstrated

- Data cleaning and validation
- Excel functions: `COUNTBLANK`, `ISBLANK`, `IF`, `LEFT`, `RIGHT`, `DATEVALUE`, concatenation (`&`)
- Find & Replace, Remove Duplicates, Spell Check
- Number and date formatting
- Conditional formatting (data bars and custom rules)
- Documenting assumptions and trade-offs

## How to Use

1. Clone or download this repository.
2. Open the `.xlsx` file in Microsoft Excel (or a compatible tool such as LibreOffice Calc or Google Sheets).
3. Review the `Instructions` sheet for the task list.
4. Compare the raw `Dataset` sheet with the cleaned `Answer` sheet.
5. Click into the helper columns in `Answer` to see the formulas used.

## Known Limitations

- **Year assumption**: Product IDs contain no year, so the manufacturing date formula assumes **2026**. Update the formula if a different year applies.
- **Currency symbol**: The column header says `($)` but the currency format applied in the workbook uses `£`. Switch to `$` for consistency.
- **Date format**: The workbook displays dates as `dd/mm/yyyy`; the brief asks for `DD-MM-YYYY`. Use a custom format of `dd-mm-yyyy` to match.
- **Missing-value labels**: `"UNKNOWN"` in the Price column turns the column into mixed text and numbers, which can break sums and averages. Consider leaving prices blank, or filtering them out before calculating.
- **Remaining gaps**: Some blank categories could be inferred from similar products rather than labelled `UNKNOWN`.

## License

This project is for educational purposes. Feel free to reuse or adapt it for your own learning.
