# Excel VBA Macro "AlphabetLoop"

#M2_VBA Scripting

#In this code, I used VBA scipting to analyze generated stock market data. 

The code initializes and declares variables to be used in the loop and finds the last row in column A of the active worksheet. It then sets the starting row for output in the "tickerRow" variable and initializes the "totalStock" variable.

This will loop through each row in column A, starting from row 2. The code will check if the value in the current row of column A is different from the previous row. If it is, it writes the ticker symbol in column 9 ("I" column) of the current "tickerRow".

The code sets the openingPrice value for calculating the yearlyChange. It checks to see if the value in the current row of column A is different from the next row. If it is, it calculates the closingPrice, yearlyChange, and percentChange based on the openingPrice. It then writes these values in columns 10, 11, and 12 respectively of the current "tickerRow".

The code updates the totalStock for the current ticker symbol. Then increments the "tickerRow" by 1 and resets the "totalStock" to 0. After the loop, it applies conditional formatting to highlight positive and negative yearly changes in column 10.
It then calculates the maximum percent change, minimum percent change, and maximum total stock using built-in Excel functions. The code will then retrieve the corresponding ticker symbols for the greatest percent increase, greatest percent decrease, and greatest total volume. It writes these values in the cells P2 to R4.

--Overall, this code loops through the data, calculates and records various metrics for each ticker symbol, applies conditional formatting, and determines the greatest percent change, percent decrease, and total volume.


--Findings:
The ticker symbol with the greatest percentage increase is "YDI" with a value of 189%.
The ticker symbol with the greatest percentage decrease is "VNG" with a value of -89%.
The ticker symbol with the greatest total volume is "QKN" with a value of 3,452,956,568,861.
