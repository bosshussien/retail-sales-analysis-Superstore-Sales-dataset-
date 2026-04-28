# Excel Sales Analysis Dashboard

## Overview
This project is an Excel-based sales analysis dashboard built from a raw transactional dataset. It includes data cleaning, calculated columns, pivot tables, and dashboards to summarize sales performance, regional trends, shipping efficiency, and profitability.

## Project Structure
- `project.xlsx` — main Excel workbook containing the full workflow
- `docs/project_report.md` — detailed project documentation
- `images/` — dashboard screenshots for preview on GitHub

## Dataset Preparation and Cleaning
Before analysis, the data was prepared with the following steps:
- Removed duplicates
- Filled missing values
- Changed data types where needed
- Created a `Profit Margin` column for later forecasting and performance analysis
- Created `Sold` and `Not Sold` columns from return-related data for aggregation
- Created a `Shipping Days` column using the difference between order date and ship date
- Grouped data by region for sold/not sold analysis

## Analysis Included
The workbook contains pivot tables and summaries for:
- Sales by time (year, quarter, and month)
- Sales by category and sub-category
- Regional order distribution
- Shipping efficiency by ship mode and region
- Profit by category and sub-category
- Regional and state-level performance
- Sold vs. not sold comparison

## Dashboard Highlights
The dashboard includes visualizations for:
- Sales trend over time
- Regional distribution of orders
- Average shipping days by region
- Shipping efficiency by ship mode
- Category profit margin comparison
- Product performance by category
- Profit by category and sub-category
- Regional/state performance overview

## Key Insights
- Technology shows the strongest profit margin among the main categories.
- Office Supplies generates the highest total sales among the three main categories.
- Standard Class has the highest average shipping time, while Same Day is the fastest.
- The West region has the highest order volume.
- Texas contributes heavily to the Central region but shows a negative profit, which is worth deeper review.
- Sales increase noticeably toward the end of 2020, especially in Q4.

## Tools Used
- Microsoft Excel
- Pivot Tables
- Charts
- Dashboard design
- Basic data cleaning and transformation

## Preview
(images)

## Notes
This project is organized as a single Excel workbook with multiple sheets:
- Raw Data
- Cleaned Data
- Pivot Tables
- Dashboard
