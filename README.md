# VBA-analysis-of-financial-data-from-green-energy-stock-performance
Analysis of financial data from green energy stock performance using VBA and Excel


## Getting Started

Import file "VBA analysis financial data.bas" into Excel project

Example Output are screen shots showing data set up and result of running VBA script


Note - there is an additonal "VBA analysis financial data bonus.bas" script that includes the above along with the following:
  greatest % increase, greatest % decrease, and greatest total volume. Will also reiterate through each sheet in Excel file


## Features

VBA script does the following:

- Loops through data to creates a summary table in columns I through J including headers

- Ticker symbol into column I

- Yearly Change (difference of openening price at beginning of year to closing price at end of year) into column J

- Conditional formatting based on value of Yearly Change (red for negative, green for positive)

- Percent Change from opening price to closing price into cloumn K (change format to percentage)

- Total Stock Volume (sum of volume per ticker symbol) into column L


## Licensing

The code in this project is licensed under MIT license.
