## Stock Market Data Analysis using VBA Scripting
The script I've created performs the following tasks:
#### Loops through all stocks for a given year and outputs the following for each respective ticker category:
1. Yearly change from the opening price to the closing price
2. Percentage change from the opening price to the closing price
3. Total stock volume
#### This script further provides functionality by finding the best of the following:
+ Stock with the greatest percentage increase
+ Stock with the greatest percentage decrease
+ Stock with the greatest total volume
***This is a second tier of analysis using the analysis of each ticker category. Stock refers to each ticker category as a whole***

### Additional features:
- Implements conditional formatting to highlight positive change in green and negative change in red
- Script is applied across all sheets within the workbook

## How the code works
The StockAnalysis subroutine is designed to analyze stock market data contained within multiple sheets in an Excel workbook using VBA scripting. Here's how it works:
1. This subroutine begins by declaring various variables, including integers for row numbers (outputRowNum, inputRowNum), a worksheet object (ws), string variables for ticker symbols (ticker, tickerPart), double variables for stock prices and volumes (openVal, closeVal, percentChange, totalStockVolume), and arrays to store the best-performing tickers and their corresponding values (allstarTickers, allstarValues).
2. It then iterates through each worksheet in the workbook. Within the loop, it initializes variables and arrays, and sets header labels for the output data.
3. During each iteration through the rows of the worksheet, it calculates the total stock volume for each ticker, determines the closing price, and prints the results including the yearly change, percentage change, and total stock volume. It also identifies the tickers with the greatest percentage increase, decrease, and total stock volume across all sheets.
4. Finally, it prints the results for each worksheet, including the top-performing tickers, and adds percentage symbols to the percentage values.

***Overall, this subroutine effectively analyzes stock market data in each worksheet, identifies top-performing stocks, and presents the results in the workbook.***

## How to Use
1. Download or clone this repository to your local machine.
2. Open the Excel workbook containing the stock market data you want to analyze.
3. Enable macros if prompted.
4. Press Alt + F8 to open the "Run Macro" dialog.
5. Select StockAnalysis() from the list and click Run.
The script will run on each sheet of the workbook, analyzing the data and displaying the results accordingly.

#### Results will look like this: 
![Screenshot 2024-02-15 at 2 43 38 PM](https://github.com/divya-jindal/VBA-challenge/assets/10901784/1c914ef7-1806-43e9-a529-72ab3cc2991c)

## Additional information
#### Files Included
* StockAnalysis().bas: The VBA script responsible for analyzing the stock market data.
* README.md: You're reading it right now! Provides information about the project.

#### Notes
Make sure to save your Excel workbook before running the script to avoid any data loss.
For any issues or suggestions, feel free to create an issue or submit a pull request.

Further, please make sure data is in columns A-G in this order: 
![image](https://github.com/divya-jindal/VBA-challenge/assets/10901784/665a8466-7fad-4965-bbbc-8246e4c215a0)

###### Credits: This project was created by Divya Jindal in UC Berkeley's Data Analytics Bootcamp taught by Instructor Kevin Lee
