# Homework #2
# Unit 2 Homework: Assignment - The VBA of Wall Street

## Description
You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, choose your assignment from Easy, Moderate, or Hard below.

# Hard

I chose the "Hard" challenge. My solution includes everything from the moderate challenge:
Created a script that will loop through all the stocks for one year for each run and take the following information: 
- The ticker symbol.
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.

I also included the conditional formatting that will highlight positive change in green and negative change in red.

My solution is also able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

# Challenge

I made the appropriate adjustments to my script that allowed it to run on every worksheet, i.e., every year, just by running it once.

# Other Considerations

I used the sheet alphabetical_testing.xlsx while developing your code. This data set is smaller and will allow you to test faster. My code ran this file in less than 1 minute.

There are additional lines of code at the bottom. I started assuming the data was not sorted alphabetically (and by date), so created a series of nested loops to solve that. It worked fine in the small dataset but crashed in the big one, after many hours solving. I needed to simplify, so I assumed all the sheets are sorted aphabetically (and by date). I reviewed the results and they look good. Thinking about it, in a real life project I would include an automatic sort at the beginning of the code to avoid miscalculations (I would need to investigate it on the internet as it was not covered in the course).


# Submission

To submit I uploaded the following to Github:


1) A screen shot for each year of your results on the Multi Year Stock Data: I decided to print the pages (one file per year) instead as a pdf file. It would be easier to check if there's any error in any of the calculations for all the tickers and all the year (Multiple_year_stock_data_NGB2*.pdf)
2) VBA Script as separate file. (HW2NEW.bas)



## Authors

* **Nicolas Gomez Bustamante** - *Initial work* - [PurpleBooth](https://github.com/nbg1)





