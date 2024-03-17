# Stock Data Scraper and Analyzer

This program is designed to scrape stock data from `stooq.pl`, a popular Polish website featuring stock quotes. It extracts stock prices for a specified ticker over the last 360 days, saves this data to an Excel file, and generates charts to visualize price changes over 90, 180, and 360 days.

## Features

- **Data Scraping**: Utilizes Selenium WebDriver to navigate `stooq.pl` and extract stock data.
- **Data Processing**: Converts the scraped data into a structured format and translates Polish date strings to English.
- **Excel Integration**: Saves the processed data into an Excel file and adds line charts for visual analysis.
- **Chart Generation**: Creates three line charts representing stock price movements over different periods (90, 180, and 360 days).

## How It Works

1. **Initialization**: The program begins by initializing the Selenium WebDriver, set to use Firefox by default.

2. **Data Scraping**: 
   - It navigates to the specific URL for the stock ticker provided by the user.
   - It handles pagination to collect up to 360 days of stock data, dealing with consent buttons as necessary.
   - The data includes the date, opening, high, low, closing price, change in percentage, nominal change, and volume.

3. **Data Processing**:
   - Translates Polish date strings into a standard "YYYY-MM-DD" format.
   - Converts the closing price to a numeric format, ensuring it can be used for calculations and charting.

4. **Excel File Creation**:
   - The program saves the structured data to an Excel file named `scraped_stock_data_[TICKER]_360_days.xlsx`.
   - It then processes this data to create a new sheet within the Excel file dedicated to charts.

5. **Chart Generation**:
   - Generates line charts for three periods: 90, 180, and 360 days.
   - Each chart displays the closing prices over its respective time period, providing a visual representation of stock price trends.

## Usage

Ensure you have the required Python libraries installed: `selenium`, `pandas`, `openpyxl`, and their dependencies.

```bash
python stock_webscraper.py PGE
