# Stock Analysis with VBA

This project contains a VBA script for analyzing stock data within Excel. The script calculates the yearly change, percent change, and total stock volume for each stock ticker. It also determines the stock with the greatest percent increase, greatest percent decrease, and greatest total volume for the year.

## Getting Started

To use this script, follow these steps:

1. Open the Excel workbook containing your stock data.
2. Press `ALT + F11` to open the Visual Basic for Applications editor.
3. Insert a new module by going to `Insert > Module`.
4. Copy the VBA script into the module window.
5. Close the editor and press `ALT + F8`, select `StockAnalysis`, and click `Run`.

## Features

- Calculation of yearly change from the opening price at the beginning of the year to the closing price at the end of the year.
- Calculation of the percentage change from the opening to the closing price over the year.
- Summation of the total stock volume for each ticker.
- Identification of the stock with the greatest percent increase, greatest percent decrease, and greatest total volume.

## Prerequisites

This script is designed to work with Excel data structured in a specific way. Ensure that each worksheet contains stock data with the following columns in order:

- Ticker Symbol
- Date
- Open Price
- High Price
- Low Price
- Close Price
- Volume

Each worksheet in the workbook should represent a year's worth of stock data.