DATA-CLEANING-AND-PREPERATION
Missing values, inconsistent formatting, duplicate records, and poorly structured text fields

Project Overview
This project is focused on cleaning, transforming, formatting, and validating a product dataset using Microsoft Excel.
Objectives

-   Identify and handle missing values
-   Standardize inconsistent text data
-   Correct category and product-name typos
-   Remove duplicate records
-   Split and merge columns using Excel formulas
-   Format dates consistently
-   Apply currency formatting to price values
-   Use conditional formatting to highlight important data
-   Prepare a clean and analysis-ready dataset

Tools & Skills Used

-   Microsoft Excel
-   Data Cleaning
-   Data Validation
-   Excel Formulas
-   Text Functions
-   Conditional Formatting
-   Duplicate Removal
-   Data Formatting
-   Data Transformation
Dataset Columns

The dataset contains product and sales-related information including:

  Column                    Description
  ------------------------- --------------------------------------
  Product ID                Unique product identifier
  Manufacturing Date 1      Original manufacturing date
  Manufacturing Date 2      Manufacturing date in another format
  Country Code              Country identifier
  Product Name              Name of the product
  Brand Name                Product brand
  Product with Brand Name   Combined product and brand
  Price                     Product price
  Quantity                  Available/sold quantity
  Category                  Product category

- Checking Missing Values in Price:
Used `IF` and `ISBLANK` to identify missing values in the Price
column.
=IF(ISBLANK(F35),AVERAGE($F$2:$F$35),F35)

- Handling Missing Values in Category:
Used `IF` and `ISBLANK` to identify missing category values.
=IF(ISBLANK(I2),"UNKNOWN",I2)
Missing category values are replaced with UNKNOWN.

- Standardizing Product Names:

Used Excel text functions to clean and standardize inconsistent
product-name formats.
=PROPER(TRIM(CLEAN(B5)))
Functions used:
-   `CLEAN()` -- removes non-printing characters
-   `TRIM()` -- removes unnecessary spaces
-   `PROPER()` -- standardizes capitalization

- Finding & Fixing Category Typos
Used Excel's **Find & Replace** feature: Ctrl + H
This was used to identify and correct inconsistent category values.

- Standardizing Product Name & Category Values
Used **Find & Replace** to fix spelling and text inconsistencies in the Product Name and Category columns.
This helps maintain consistent categorical values for further analysis.

- Removing Duplicate Rows
Used Excel's Remove Duplicates feature to identify and remove duplicate records from the dataset.
Select the dataset → Data → Remove Duplicates → Select all columns → Remove Duplicates
This ensures that each record is represented only once.

- Splitting Product ID into Manufacturing ID & Country Code
The Product ID was split into two separate fields:
-   **Manufacturing ID**
-   **Country Code**
Example:
Product ID: LA7,6
Manufacturing ID:
=LEFT(A7,6)

Country Code:
=RIGHT(A3,2)
This makes the identifier easier to analyze.

- Merging Product Name & Brand Name Combined Product Name and Brand Name into a single field.
=CONCATENATE(F3," ",G3)
Example:
LAPTOP + DELL→ LAPTOP DELL

- Formatting Price as Currency
Applied **Indian Rupee (₹)** currency formatting to the Price column.
Select Price column → Home → Number → Currency → ₹
This improves readability and ensures consistent financial formatting.

- Formatting Manufacturing Dates
Used Excel's date formatting to standardize manufacturing dates.
Formula used for conversion: =DATEVALUE(B3)
The resulting values were formatted consistently as dates.

- Highlighting Electronics Category
Applied **Conditional Formatting** to highlight products belonging to
the **Electronics** category.
Select Category column → Home → Conditional Formatting → Highlight Cells Rules → Text that Contains → Electronics
This makes the target category easy to identify visually.

- Applying Color Scales to Price
Applied **Conditional Formatting → Color Scales** to the Price column.
This provides a quick visual representation of lower and higher product prices and makes patterns easier to identify.

- Key Data Cleaning Techniques Demonstrated

  Technique                Excel Feature / Formula
  ------------------------ ----------------------------
  Missing value handling   `IF`, `ISBLANK`, `AVERAGE`
  Text cleaning            `CLEAN`, `TRIM`, `PROPER`
  Typo correction          Find & Replace
  Duplicate removal        Remove Duplicates
  Column splitting         `LEFT`, `RIGHT`
  Column merging           `CONCATENATE`
  Date conversion          `DATEVALUE`
  Currency formatting      Number → Currency
  Category highlighting    Conditional Formatting
  Price visualization      Color Scales

- Project Outcome

The raw product dataset was transformed into a cleaner, standardized, and analysis-ready dataset by applying multiple Excel data-cleaning and transformation techniques.

This project demonstrates my foundational skills in Excel-based data analytics and data preparation.
