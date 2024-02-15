# VBA-challenge

## Stock Market Data Analysis using VBA Scripting
The script I've created performs the following tasks:

### Loops through all stocks for a given year and outputs the following for each respective ticker category:
1. Yearly change from the opening price to the closing price
2. Percentage change from the opening price to the closing price
3. Total stock volume

### This script further provides functionality by finding the best of the following:
+ Stock with the greatest percentage increase
+ Stock with the greatest percentage decrease
+ Stock with the greatest total volume
***This is a second tier of analysis using the analysis of each ticker category. Stock refers to each ticker category as a whole***

##### Additional features:
##### - Implements conditional formatting to highlight positive change in green and negative change in red
##### - Script is applied across all sheets within the workbook


## How to Use
1. Download or clone this repository to your local machine.
2. Open the Excel workbook containing the stock market data you want to analyze.
3. Enable macros if prompted.
4. Press Alt + F8 to open the "Run Macro" dialog.
5. Select StockAnalysis() from the list and click Run.
The script will run on each sheet of the workbook, analyzing the data and displaying the results accordingly.

#### Files Included
* StockAnalysis().bas: The VBA script responsible for analyzing the stock market data.
* README.md: You're reading it right now! Provides information about the project.

#### Notes
Make sure to save your Excel workbook before running the script to avoid any data loss.
For any issues or suggestions, feel free to create an issue or submit a pull request.

###### Credits
###### This project was created by Divya Jindal
