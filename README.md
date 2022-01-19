# VBA of Wall Street

## Overview of Project

This project is an analysis of a selection of stock market data from the years 2017 and 2018. We use Microsoft Excel and VBA scripts to analyze the data from twelve different companies and calculate the total daily volume traded and the yearly return.

### Purpose

We are helping Steve automate the analysis of data on stocks he has chosen. His parents are currently only invested in one stock and he wants to help choose more for them. In addition, it is important that the script runs quickly and efficiently so that Steve can run it on any stock data that he wants to.

## Analysis

### Stock Perfomance

After performing the analysis for the years 2017 and 2018, it's immediately apparent that in 2018 almost all stocks had a negative return on investment, while in 2017 almost all stocks had a positive return on investment. Only **ENPH** and **RUN** had positive returns for both years. The stocks of *SEDG* and *VSLR* had good returns in 2017 and only slightly negative returns in 2018. Without further analysis and comparison it's hard to make any other conclusions, but I would recommend those four stocks for Steve's parents. 

### Code performance

The refactored script ran approximately eight times faster than the original script, as seen in the screenshots below. 
>![Runtime for original code on 2017 data](resources/VBA_Original_2017.png) ![Runtime for original code on 2018 data](resources/VBA_Original_2018.png)\
>Runtimes for the orignal script.


>![Runtime for refactored code on 2017 data](resources/VBA_Challenge_2017.png) ![Runtime for refactored code on 2018 data](resources/VBA_Challenge_2018.png)\
>Runtimes for the refactored script.

This increase in performance is because the refectored script only runs through every row once, while the original runs through every row twelve times. While the refactored script has more arrays and needs to access them more often, it is more than made up for by the number of times it has to loop through the rows. 

## Summary

In summary, the VBA scripts are a powerful tool for working with data in Microsoft Excel, even though they are a big security risk. Refactoring the script was able to get a large increase in efficiency.

### Advantages and Disadvantages of Refactoring

When done properly, refactoring allows old code to be made more efficient, whether that means a faster execution time, less memory usage, or some other metric. However, in addition to the time needed to learn the original code, and come up with a new version, refactoring can also lead to code is more complicated or less intuitive to follow. Good use of comments and documentation should help with this, but it's still a possibility.

### Advantages and Disadvantages of our VBA Script

In our case, refactoring resulted in a large decrease in the runtime of our script. Looking at our code, we can see that in the original we have to run through every row twelve times in order to collect the data for each stock.
```
For j = 0 To 11
        ticker = tickers(j)

        totalVolume = 0
        Worksheets(yearValue).Activate

        For i = rowStart To rowEnd
```
For our refactored script we only have to run through all the rows twice.
```
For i = 2 To RowCount
```

One downside of our refactored code is that it depends on the tickers in the worksheet being in the same order as they are in our tickers array. Since we have `tickerIndex = tickerIndex + 1` then a ticker that is out of order compared to the array would be skipped over by all of the `If` Statements in the code.
