# Data Modeling and Analysis with Excel Power Pivot and Power Query


This Excel project demonstrates data modeling and analysis using Power Pivot and Power Query. Three datasets (orders, products, and customers) were connected and analyzed to gain insights into sales and returning customers. Power Query was used for minimal data cleaning, while Power Pivot facilitated data modeling through relationships and calculated columns. Analysis included creating a profit heatmap by coffee and roast type, visualizing profit trends with a line chart, and identifying returning customers with the highest order counts and profit per order. The project highlights the ability to efficiently model and analyze data from multiple sources within Excel, showcasing skills in data cleaning, modeling, DAX calculations, and visualization.

https://github.com/user-attachments/assets/a6e18d14-3cf1-4355-b8c5-db18222160fa

## Process Breakdown
**1. Data Acquisition and Preparation (Power Query):**
 - Imported data from three CSV files: orders, customers, and products.
 - Used Power Query to connect to these files and load the data.
 - Performed minimal data cleaning within Power Query (no significant transformations were mentioned).
 - Loaded the data into the Power Pivot data model, creating connections only at this stage.
**2. Data Modeling (Power Pivot):**
 - Opened the Power Pivot window and created a data model.
 - Established relationships between the tables:

orders and customers (one-to-many, customer_id)
orders and products (one-to-many, product_id)


 - Generated a date table using Power Pivot's built-in date table generation feature (Design > Date Table > New).
 - Connected the orders table to the generated date table using the order_date and date fields.
 - Created calculated columns in the orders table:

profit: Calculated by multiplying quantity (from orders) with profit_per_unit (from products, using the RELATED() function).
returning_customer_flag:  Used a CALCULATE() and COUNTROWS() combination to determine if a customer ID appeared more than once in the orders table.


**3. Data Analysis and Visualization (Excel):**
 - Inserted a pivot table into a new worksheet.
 - Created a profit heatmap:

Placed coffee_type and roast_type in the rows.
Placed profit in the values.
Used conditional formatting (color scales) to visualize profit levels.
Added a date hierarchy (years, months) to the pivot table for drill-down analysis.


 - Created a line chart to visualize profit trends:

Moved coffee_type to the legend.
Used the date hierarchy (years, months) on the axis.
Placed profit in the values.
Customized the chart's appearance (layout, colors).


 - Created a pivot table to analyze returning customers:

Filtered the returning_customer_flag column to show only "True" values.
Placed customer_name and email in the rows.
Placed order_id (count) in the values.
Sorted customers by the count of order IDs.


 - Created a measure for profit per order in Power Pivot:

profit_per_order: Calculated by dividing SUM(profit) by COUNT(order_id).


Added the profit_per_order measure to the returning customers pivot table.
Sorted customers by profit_per_order (descending).
Used slicers (e.g., for roast_type) to interactively filter the data.
