# Stock-Analysis

## Overview of Project

Our client, Steve, likes the analysis we did on the stock market, but he wants to run our code on a larger data set. We need to refactor our code to make everything run faster, making it easier for Steve to do analysis on larger datasets. We will compare our run times for the original code with the refactored code, to see if the refactored code is successful.


### Purpose

The purpose of this analysis is to edit our old code to make it faster, making it easier to run more data, quickly. We will then compare our new numbers to the old, to see if the refactoring made a difference.

## Results

In order to make our code better for larger datasets, we edited it to use fewer steps, and less memory, making our code more efficient. The old code iterated over each row for each of the 12 tickers. The new code only iterates over each row once, and determines which ticker to assign the data to, which means less iterations and faster. The old method in pseudocode is below, which shows the row loop nested in the ticker loop.

``` 
    ' Loop through the tickers
    For i = 0 To 11
        ' Loop through rows in the data
        For j = 2 To RowCount
            ' Find total volume for the current ticker
            
            ' Find starting price for the current ticker
            
            ' Find ending price for the current ticker
        Next j
        
    ' Output the data for the current ticker
    Next i
```
The new method in pseudocode, below, shows the row loop is not nested in the ticker loop making it run faster.

```
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount

        '3a) Increase volume for current ticker

        '3b) Check if the current row is the first row with the selected tickerIndex.
        ' set the starting price
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.

    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For k = 0 To 11
    Next k
```
The old code is located in Module 1 of the VBA_Challenge.xlsm workbook (yearValueAnalysis), and the new code is located in Module 2 of the VBA_Challenge.xlsm workbook (AllStocksAnalysisRefactored).

### 2017 Results

The amount of time it originally took our code to run the data for 2017 (before refactoring) was about 0.623 seconds, as seen in the image below.

![image info](./resources/VBA_Challenge_2017_not_refactored.png)

The amount of time it took our code to run the data after refactoring was about 0.0957 seconds, as seen in the image below.

![image info](./resources/VBA_Challenge_2017.png)

Clearly the new algorithm runs faster because of fewer iterations.

### 2018 Results

The amount of time it originally took our code to run the data for 2018 (before refactoring) was about 0.506 seconds, as seen in the image below.

![image info](./resources/VBA_Challenge_2018_not_refactored.png)

The amount of time it took our code to run the data after refactoring was about 0.0801 seconds, as seen in the image below.

![image info](./resources/VBA_Challenge_2018.png)

We can see that the new time is faster than the old time with the refactored code.

## Summary

The biggest advantage of refactoring code is having a faster run time, which is helpful for large datasets. With refactored code, less memory is used and makes the code run faster. The larger the dataset gets, the bigger a difference you will see. The biggest disadvantage of refactoring code is that it takes time to change code that already works. It takes time to analyze the code and figure out where improvements can be made. This applies to our stock analysis because we already had code that ran fairly quickly (under one second). The new code clearly runs much faster than the old code. However it is questionable whether the time spent refactoring the code was worth the effort, since our dataset isn't very large.

