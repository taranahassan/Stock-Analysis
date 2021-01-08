# stock-analysis
## Performing Green Energy analysis to make informed decisions for investments

Steve being a recent graduate with a Finance degree, his parents approached him to assist them with their investment.  His parents are passionate about the Green Energy sector.  They believe fossil fuels will eventually deplete, therefore will be more dependant on alternate energy.  There are four components to the Green Energy sector; Hydro, Geothermal, Wind and Bio.  However Steve's parents have not done alot of research and therefore wanted to invest in DAQO New Energy Corp. who manufactures silicon waffels for solar panels.

Steve, as their financial advisor is hesitant and would like to diversify his parents' investment by analyzing deeper into all stocks for energy companies including DAQO's stocks to make a more informed decision.

### Overview of Project

This project is to analyze all the stock data that has been compiled by Steve.  VBA was used to automate the task of analyzing which is more efficient when dealing with larger datasets.  Steve would be able to compare data from each year with regards to the stock prices, total daily volume and return for each stock with a click of a button.  The worksheet has been saved as an Excel Macro-enabled for Steve to access.  

### Results

When comparing 2017 stocks with 2018 stocks, most companies had a huge dip in their return year over year.  For 2017, all stocks had a positive return except a company named TERP.  Since Steve's parents were interested in DAQO, this company had the second highest return in 2017; just a little below 200%.

For 2018, all stocks took a hit in their return except ENPH and RUN.  DAQO, being a company of interest, had a negative return of 62%.  Though ENPH had a positive return for both years, they dropped about 38% from 2018.  

Out of all the stocks, RUN had a growth return of 78.5% year over year.  Their total daily volume traded has almost doubled as well within the year.

Screenshots of data for both 2017 and 2018 are included.  [2017_stock_analysis](https://github.com/taranahassan/stock-analysis/blob/main/2017_stock_analysis.png?raw=true).  [2018_stock_analysis](https://github.com/taranahassan/stock-analysis/blob/main/2018_stock_analysis.png?raw=true)


Initially to create the code to run the analysis task, it was a bit slower to execute.  Both years took about 1.60 seconds to execute on the worksheet.  Screenshots of execution times for original code included.  [green_stocks_2017](https://github.com/taranahassan/stock-analysis/blob/main/green_stocks_2017.png?raw=true).  [green_stocks_2018](https://github.com/taranahassan/stock-analysis/blob/main/green_stocks_2018.png?raw=true).

After refactoring code the executions times reduced almost by half.  Both 2017 and 2018 analysis executed within 0.17 seconds.  [VBA_Challenge_2017](https://github.com/taranahassan/stock-analysis/blob/main/VBA_Challenge_2017.png?raw=true).  [VBA_Challenge_2018](https://github.com/taranahassan/stock-analysis/blob/main/VBA_Challenge_2018.png?raw=true).
