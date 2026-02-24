# Regression Project

This project performs linear regression analysis on 2024 data from Raw Data.xlsx to predict Revenue based on various features.

## Setup

1. Ensure you have Python installed.
2. Install dependencies: `pip install -r requirements.txt`
3. Place your data file as `Raw Data.xlsx` in the project directory (sample outputs are included for reference).
4. Run the script: `python regression.py`

## Data

The data should be in Excel format with columns for Year, Quarter, Product Model, etc. The script filters for 2024 data. Features include Units Sold, Market Share, etc., and the target is Revenue ($).

## Output

The script prints a summary, saves a detailed Excel report to `regression_report.xlsx`, and a chart to `regression_plot.png`. Open the Excel file for a professional, tabular view with multiple sheets.

## Email Feature

To send the report and chart via email:
1. Set environment variables: `export EMAIL_SENDER='your_email@gmail.com'` and `export EMAIL_PASSWORD='your_app_password'`
2. Run: `python send_email.py`

Use a Gmail app password for EMAIL_PASSWORD (not your regular password). Enable "Less secure app access" if needed.

## Troubleshooting

- Ensure Raw Data.xlsx is in the same directory.
- For email, use a Gmail app password and update sender details in `send_email.py`.# Initial-Regression-Analysis-Project
