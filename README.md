# stock-analysis
## Performing Green Energy analysis to make informed decisions for investments

### Overview of Project

Steve being a recent graduate with a Finance degree, his parents approached him to assist them with their investment.  His parents are passionate about the Green Energy sector.  They believe fossil fuels will eventually deplete, therefore will be more dependant on alternate energy.  There are four components to the Green Energy sector; Hydro, Geothermal, Wind and Bio.  However Steve's parents have not done alot of research and therefore wanted to invest in **DAQO New Energy Corp.** who manufactures silicon waffels for solar panels.
Steve, as their financial advisor is hesitant and would like to diversify his parents' investment by analyzing deeper into all stocks for energy companies including **DAQO's** stocks to make a more informed decision.  
This project is to analyze all the stock data that has been compiled by Steve.  VBA was used to automate the task of analyzing which is more efficient when dealing with larger datasets.  Steve would be able to compare data from each year with regards to the stock prices, total daily volume and return for each stock with a click of a button.  If in the future Steve includes another sheet of data for another year, the code would be able to calculate that as well.  The worksheet has been saved as an Excel Macro-enabled for Steve to access.  


### Results

When comparing 2017 stocks with 2018 stocks, most companies had a huge dip in their return year over year.  For 2017, all stocks had a positive return except a company named **TERP**.  Since Steve's parents were interested in **DAQO**, this company had the second highest return in 2017; just a little below 200%.
For 2018, all stocks took a hit in their return except **ENPH** and **RUN**.  **DAQO**, being a company of interest, had a negative return of 62%.  Though **ENPH** had a positive return for both years, they dropped about 38% from 2017.  
Out of all the stocks, **RUN** had a growth return of 78.5% year over year.  Their total daily volume traded has almost doubled as well within the year.

*Screenshots of data for both 2017 and 2018 are included.  
![2017_stock_analysis](https://github.com/taranahassan/stock-analysis/blob/main/2017_stock_analysis.png?raw=true).  ![2018_stock_analysis](https://github.com/taranahassan/stock-analysis/blob/main/2018_stock_analysis.png?raw=true)* 

Initially creating the code to run the analysis task was a bit slower to execute.  Both years took about **1.60 seconds** to execute on the worksheet.  

*Screenshots of execution times for original code included.  
![green_stocks_2017](https://github.com/taranahassan/stock-analysis/blob/main/green_stocks_2017.png?raw=true).  ![green_stocks_2018](https://github.com/taranahassan/stock-analysis/blob/main/green_stocks_2018.png?raw=true).*  

This is because nested loops were incorporated in the original VBA script.  
![Original_VBA_script](https://github.com/taranahassan/stock-analysis/blob/main/Original_VBA_script.png?raw=true)

I believe the nested loop in the code had to circle through twice to execute the results.

After refactoring the code, executions times reduced almost more than half.  Both 2017 and 2018 analysis executed within **0.17 seconds**.  
*![VBA_Challenge_2017](https://github.com/taranahassan/stock-analysis/blob/main/VBA_Challenge_2017.png?raw=true).  ![VBA_Challenge_2018](https://github.com/taranahassan/stock-analysis/blob/main/VBA_Challenge_2018.png?raw=true).*  

The nested loop was taken out and single loops were created for each section of the data.
![Refactored_VBA_script](https://github.com/taranahassan/stock-analysis/blob/main/Refactored_VBA_script.png?raw=true)


### Summary

#### *_Advantage vs. Disadvantage_*
The advantages of refactoring a code is making the macro run more efficiently.  This reduces the time to execute an analysis to get results faster and possibly with less errors.  This is beneficial especially with larger datasets.  However if you don't understand the code well enough, or don't understand the properties of the code (such as refactoring someone else's code), then refactoring the code could end up with the macro delaying the process further or even not function at all.  By refactoring you can also clean and organize the code.  If not within a time frame, refactoring a code can be properly restructed if necessary.
Refactoring the code is ideal when you want to create less nested loops.  Nested loops will loop over the data twice as stated.  But when the code is refactored, there is a possibility to create single loops for each part of the data as long as each variable is stated and linked.  However this might not be avoidable with datasets that require more intricate analysis.

#### *_Pros & Cons_*
When refactoring the original VBA script, the execution was much faster when analyzing each year.  However it got confusing when trying to link the ticker index to each variable that was created.  Again, this could lead to the code not working at all!  But in this case it did and the macro ran much smoother with less time plus the script size was minimized.
