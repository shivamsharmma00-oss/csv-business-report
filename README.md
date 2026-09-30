# Automated Business Sales Reporting System

A Python-based business reporting tool that cleans messy sales CSV data, performs sales analysis, identifies key business insights, and automatically generates a structured Excel report.

## Features

- Validates required CSV columns
- Removes duplicate records
- Handles missing and invalid sales amounts
- Cleans currency-formatted values such as `₹55,000`
- Normalizes product and salesperson names
- Handles missing product and salesperson values
- Calculates total sales
- Calculates average sale
- Calculates median sale
- Finds the highest and lowest sale
- Counts the number of transactions
- Analyzes sales performance by product
- Analyzes sales performance by salesperson
- Identifies the top-performing product
- Identifies the top-performing salesperson
- Generates an automated Excel business report
- Creates separate sheets for summary, product analysis, salesperson analysis, and cleaned data
- Adds formatting and sales visualizations to the Excel report

## Tech Stack

- Python
- Pandas
- OpenPyXL

## How to Run

### 1. Clone the repository

~~~bash
git clone https://github.com/shivamsharmma00-oss/csv-business-report.git
cd csv-business-report
~~~

### 2. Install dependencies

~~~bash
pip install pandas openpyxl
~~~

### 3. Run the program

~~~bash
python main.py
~~~

The program will ask for the CSV filename:

~~~text
Enter CSV filename:
~~~

For example:

~~~text
sales.csv
~~~

The program will then clean the data, perform the analysis, display a business report in the terminal, and generate:

~~~text
sales_business_report.xlsx
~~~

## Input Format

The program expects a CSV file containing these required columns:

~~~csv
product,salesperson,amount
Laptop,Rahul,55000
Mouse,Priya,2500
Keyboard,Aman,4000
~~~

The program validates these required columns before processing the data.

## Data Cleaning

The pipeline handles common problems found in real-world CSV data, including:

- Duplicate records
- Missing values
- Invalid numeric values
- Currency symbols
- Commas in currency values
- Extra whitespace
- Inconsistent capitalization
- Missing product or salesperson names

Invalid sales amounts are converted to missing values and excluded from the financial calculations.

## Generated Excel Report

The generated Excel workbook contains four sheets.

### Summary

Contains key business metrics and automatically calculated business insights.

### Sales by Product

Contains:

- Total sales
- Average sale
- Number of transactions

for each product.

### Sales by Salesperson

Contains:

- Total sales
- Average sale
- Number of transactions

for each salesperson.

### Cleaned Data

Contains the cleaned dataset used for the final analysis.

## Project Structure

~~~text
csv-business-report/
├── main.py
├── sales.csv
├── README.md
└── .gitignore
~~~

## Example Use Case

This project demonstrates how a business can automate a repetitive reporting workflow:

~~~text
Raw CSV Data
     ↓
Validation
     ↓
Data Cleaning
     ↓
Business Analysis
     ↓
Business Insights
     ↓
Automated Excel Report
~~~

Instead of manually cleaning sales data and preparing reports, the entire workflow can be executed with a single Python program.