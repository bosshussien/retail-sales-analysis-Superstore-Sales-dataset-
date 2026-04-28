# Project Report: Excel Sales Analysis Dashboard

## 1. Introduction
This project analyzes sales performance using Microsoft Excel. The goal is to transform raw transactional data into a clear and interactive dashboard that can be used to understand sales, shipping behavior, profitability, and regional performance.

## 2. Project Objective
The main objective of this project is to:
- Clean and prepare raw data
- Build useful summary tables
- Visualize performance trends
- Identify areas of strong and weak performance
- Support future forecasting and decision-making

## 3. Data Cleaning Process
The raw dataset was prepared using basic cleaning and transformation steps:
- Duplicate records were removed
- Missing values were filled
- Data types were corrected where necessary
- A `Profit Margin` column was created for later forecasting and comparison
- `Sold` and `Not Sold` columns were created from return information for grouped analysis
- A `Shipping Days` column was created by calculating the difference between order date and ship date
- Region-based aggregation was added to compare sold and not sold counts

These steps helped make the data more consistent and suitable for analysis.

## 4. Pivot Tables and Analysis
Several pivot tables were created to summarize the data from different perspectives.

### 4.1 Sales by Time
A time-based pivot table was created to analyze sales by year, quarter, and month.  
This made it possible to see how sales changed over time and to identify seasonal patterns.

### 4.2 Product Performance
A product performance pivot table was built by category and sub-category.  
This was used to compare:
- Total sales
- Total profit
- Average profit margin

### 4.3 Shipping Efficiency
Shipping performance was analyzed by:
- Ship mode
- Region
- Average shipping days

This helped compare how fast orders were delivered across different shipping methods and regions.

### 4.4 Regional Analysis
A regional pivot table was created to compare:
- Sold
- Not sold
- Grand total

A second regional and state-level pivot table was also created to examine:
- Shipping days
- Quantity
- Profit

This allowed a deeper look at performance by location.

## 5. Dashboard Description
The Excel dashboard combines the main visuals in one place for easier interpretation.

### Included Visuals
- Sales trend by time
- Average shipping days by region
- Regional distribution of orders
- Category profit margin comparison
- Shipping efficiency by ship mode
- Product performance chart
- Profit by category
- Regional/state performance chart

The dashboard gives a quick overview of the most important business metrics.

## 6. Main Findings
Based on the dashboard and pivot tables, the following insights were observed:
- Technology has the highest average profit margin among the main categories
- Office Supplies has the highest total sales
- Standard Class takes the longest shipping time on average
- Same Day shipping is the fastest
- The West region has the highest order volume
- Texas shows a negative profit in the Central region, which suggests a possible issue that needs review
- Sales increased strongly in Q4 of 2020 compared with earlier periods

## 7. Conclusion
This project demonstrates how raw Excel data can be cleaned, transformed, analyzed, and presented in a dashboard format. It shows a complete workflow from data preparation to insight generation and is suitable for portfolio presentation on GitHub.

## 8. Tools Used
- Microsoft Excel
- Pivot Tables
- Charts and dashboards
- Basic data cleaning and transformation
