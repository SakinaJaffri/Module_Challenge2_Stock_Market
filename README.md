# Stock Market Analysis using VBA
![image](https://github.com/SakinaJaffri/Module_Challenge2_Stock_Market/assets/146900226/8d518f95-0d74-478b-be97-34baae38fb60)

## Project Overview
This project utilizes VBA scripting to analyze stock market data for multiple years. The script loops through all the stocks for a given year, performing calculations on yearly change, percentage change, and total stock volume, while identifying the stocks with the greatest increase, decrease, and volume.

## Features
- **Ticker Symbol**: Extracts and displays the stock ticker symbol.
- **Yearly Change**: Calculates the yearly change in stock price from the opening to the closing of the year.
- **Percentage Change**: Calculates the percentage change in stock price from the opening to the closing of the year.
- **Total Stock Volume**: Sums the total volume of stocks traded during the year.
- **Greatest Values**: Identifies the stock with:
  - Greatest % increase
  - Greatest % decrease
  - Greatest total volume

## Code Summary
- **Column Creation**: Ticker, Yearly Change, Percent Change, and Total Stock Volume.
- **Calculated Values**: Greatest % increase, Greatest % decrease, and Greatest Total Volume.
- **Looping Function**: VBA script loops across all worksheets (years) in the workbook, analyzing each year's data.
- **Conditional Formatting**: Applied to the Yearly Change and Percent Change columns to visually differentiate performance.
- **Max/Min Calculation**: Populates and highlights the maximum and minimum values for percentage changes and total volume.
- **Color Indexing**: Applied to cells based on performance metrics (e.g., positive/negative changes).

## VBA Script Highlights
- Loops through stock data for each year.
- Assigns calculated values to new columns.
- Runs on all sheets at once using `For Each ws In Worksheets`.

## Screenshots
The repository includes screenshots for the stock data from 2018, 2019, and 2020 for better understanding of the VBA script's output.

## Tools Used
- Microsoft Excel
- VBA (Visual Basic for Applications)

## Contributors
- **Sakina Jaffri** - VBA scripting, data analysis, and report creation.
