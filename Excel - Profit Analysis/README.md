# A Practical Excel Analysis of Gym Equipment Profitability

Project Overview:

 This project demonstrates a practical approach to data analysis using Microsoft Excel, focusing on extracting actionable insights to solve real-world business problems.  Using a fictitious gym equipment profit dataset, I performed a comprehensive analysis covering key areas such as supplier/brand/category performance, year-on-year growth trends, market share dynamics, and year-to-date (YTD) and moving annual total (MAT) profit calculations.  The project emphasizes efficiency and clarity, prioritizing insightful interpretations over complex visualizations.  My analysis reveals crucial performance indicators, enabling data-driven decision-making for optimizing profitability.

https://github.com/user-attachments/assets/9040550c-081c-4b93-964b-4bd363f06aa3

## Process Breakdown
**1. Data Preparation and Understanding:**
Inspect Data: Examined the dataset to understand the columns (category, supplier, brand, year, month, monthly profit) and their data types.
Tidy Data: Used double-clicking on column headers to auto-adjust column widths for better readability.
Convert to Table: Transformed the data range into an Excel Table (using Ctrl+T) to enable dynamic updates for pivot tables and charts.

**2. Basic Descriptive Analysis:**
Create Pivot Table: Inserted a pivot table (Alt+NVT or Insert > PivotTable) to start the analysis.
Supplier/Brand/Category Counts: Dragged the respective fields (supplier, brand, category) into the pivot table to determine the number of unique entries for each. Used tabular form layout (Design > Report Layout > Show in Tabular Form) for better presentation.

**3. Year-on-Year Growth Analysis:**
Duplicate Worksheet: Duplicated the pivot table worksheet (Ctrl+drag) to preserve the original pivot table setup.
Calculate YoY Growth:  Moved "Year" to the Columns area and "Monthly Profit" to the Values area.  Used Value Field Settings > Show Values As > % Difference From to calculate year-on-year growth.  Set the base field to "Year" and base item to "Previous."
Conditional Formatting: Applied conditional formatting (Home > Conditional Formatting > Color Scales) to visually highlight growth and decline percentages. Removed 2024 data due to incomplete data for that year.

**4. Market Share Analysis:**
Duplicate Worksheet: Duplicated the worksheet again.
Calculate Market Share: Moved "Year" to the Rows area, "Brand" to the Columns area, and "Monthly Profit" to the Values area. Used Value Field Settings > Show Values As > % of Row Total to calculate market share for each brand within each year.
Insert Slicer: Added a slicer (PivotTable Analyze > Insert Slicer) for "Category" to filter market share by product category.
Insert Chart: Created a line chart to visualize the market share trends over time.  Modified the chart type and formatting for better visualization.

**5. Year-to-Date (YTD) and Moving Annual Total (MAT) Profit Calculations:**
Add Calculated Columns: Added two new columns to the original data table: "Profit Year to Date" and "Profit Moving Annual Total."
Calculate YTD Profit: Used SUMIFS formulas to calculate the cumulative profit for each month, considering the category, supplier, brand, year, and month. The SUMIFS formula summed profits from the beginning of the year up to the current month.  Formatted the column as currency.
Calculate MAT Profit: Used SUMIFS formulas combined with the "Profit Year to Date" value to calculate the total profit for the trailing 12 months.  This involved summing the YTD profit of the current year and the profit from the corresponding months of the previous year to complete the 12-month period. Formatted the column as currency.
Sense Checks: Performed manual calculations to verify the accuracy of the YTD and MAT profit calculations.

**6. Pivot Table Analysis with Calculated Fields:**
Duplicate Worksheet: Duplicated the worksheet again.
Refresh Pivot Table: Refreshed the pivot table (PivotTable Analyze > Refresh All) to include the newly calculated columns.
Pivot Table Setup: Removed "Monthly Profit," moved "Year" and "Month" to the Filters area, "Brand" and "Supplier" to the Rows area.
Filter Data: Filtered the pivot table to show data for the last month of the available data (May 2024).
Format Values: Formatted the values in the pivot table as currency.

## Insights from the selected sheets:

1. Profit SheetThe data appears to be profit values per supplier and brand across different years (2018-2023).
Some profit values are negative, indicating potential losses in certain years.Example:Forge Fitness (Iron Strength Equipment Co.) had a loss in 2019 (-0.0317) but a profit in 2020 (0.0508).Ironclad Athletics had fluctuating values, with losses in 2019, 2020, and 2022, but gains in 2021 and 2023.2. 

Line Chart Sheet
This sheet likely supports a visual representation of profit trends.
It contains profit values per brand over the years (2018-2019 shown).
The brands included:Apex Athletics, Elevate Fitness, Spartan Sports, Steel Power, Summit Strength, and Titan Training.
The values indicate profitability trends, where most brands had consistent profits with slight variations.

3. MAT by Month Sheet
This sheet contains Moving Annual Total (MAT) profits.
The data includes:Supplier, Brand, and Product Category (e.g., Rowing Machine).Sum of Profit MAT for 2024 (e.g., Forge Fitness: 150,568).
This is useful for tracking annual performance trends.

## Summary
This project showcases my ability to leverage Excel's powerful features to conduct meaningful data analysis.  By focusing on practical problem-solving and clear communication of insights, I've demonstrated how data can be transformed into actionable knowledge for driving business success.

