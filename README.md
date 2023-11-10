# Module_Challenge2_Stock_Market
BACKGROUND:
![image](https://github.com/SakinaJaffri/Module_Challenge2_Stock_Market/assets/146900226/8d518f95-0d74-478b-be97-34baae38fb60)


In this assignment, we used VBA scripting to analyze generated stock market data.

Created a script that loops through all the stocks for one year and outputs the following information:

•	The ticker symbol

•	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

•	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

•	The total stock volume of the stock.

•	Added functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"

•	Made appropriate adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once.





Stock Market Analysis Code/Script (via these indictors):

Retrieval of Data

•	The script loops through one year of stock data and reads/ stores all of the following values from each row:

o	ticker symbol

o	volume of stock 

o	open price 

o	close price 

Column Creation

•	On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:

o	ticker symbol 

o	total stock volume 

o	yearly change ($) 

o	percent change 

Conditional Formatting 

•	Conditional formatting is applied correctly and appropriately to the yearly change column 

•	Conditional formatting is applied correctly and appropriately to the percent change column 

Calculated Values 

•	All three of the following values are calculated correctly and displayed in the output:

o	Greatest % Increase

o	Greatest % Decrease

o	Greatest Total Volume 

Looping Across Worksheet

•	The VBA script can run on all sheets successfully.




Code Summary:

Created Ticker, Yearly Change, Percent Change & Total Stock Volume (Columns)

Assigned the Ticker, Value (columns) & Greatest % increase, Greatest % decrease, Greatest Total volume (values in the columns), respectively!

Looping function & Vlookup function performed to populate the Ticker, Yearly change, Percent Change & Total Stock Volume respectively!

Color Index in the respective columns!

Populated and matched the Max, Min and Total volume (values) in the respective columns!

Made appropriate adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once through "For Each ws. In Worksheets"



Note: please find the attached 3 screenshots of Excel Sheets (2018,2019 & 2020) in the respository for better understanding dataset scripted through using VBA. 

