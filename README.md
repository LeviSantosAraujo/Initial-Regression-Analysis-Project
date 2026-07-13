# Regression Project

This project performs linear regression analysis on 2024 data from Raw Data.xlsx to predict Revenue based on various features.

## Setup

1. Ensure you have Python installed.
2. Install dependencies: `pip install -r requirements.txt`
3. Place your data file as `Raw Data.xlsx` in the project directory (sample outputs are included for reference).
4. Run the script: `python regression.py`

## Data

The data should be in Excel format with a Year column for filtering. The script filters for 2024 data and uses the following features:
- Units Sold
- Market Share (%)
- Regional 5G Coverage (%)
- 5G Subscribers (millions)
- Avg 5G Speed (Mbps)
- Preference for 5G (%)

The target variable is Revenue ($).

## Output

The script prints a summary, saves a detailed Excel report to `regression_report.xlsx`, and a chart to `regression_plot.png`. Open the Excel file for a professional, tabular view with multiple sheets.

## Email Feature

To send the report and chart via email:
1. Enable 2-factor authentication on your Gmail account
2. Generate an App Password at https://myaccount.google.com/apppasswords
3. Set environment variables: `export EMAIL_SENDER='your_email@gmail.com'` and `export EMAIL_PASSWORD='your_app_password'`
4. Run: `python send_email.py`

Use a Gmail app password for EMAIL_PASSWORD (not your regular password).

## Troubleshooting

- Ensure Raw Data.xlsx is in the same directory.
- For email, use a Gmail app password and update sender details in `send_email.py`.# Initial-Regression-Analysis-Project
