
# VBA Challenge

## Overview of Project

The stock analysis determines the total volume traded and gain or loss over the period of a year. Our first approach iterated over the full year's data once for each stock ticker. While this approach works, it is not scalable if the number of tickers is significantly increased.

This analysis is intended to determine if the performance of the stock analysis code can be improved by refactoring such that only one iteration is needed over the source data.

## Results

In the initial approach, analyzing 2017 data took approximately 1.6 seconds and 2018 took approx. 1.7 seconds.

The code contains a nested loop which repeats the iteration across all rows in the target worksheet:

```VBA
   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           ...
       Next j
       ...
    Next i
```

By using arrays to store the output we can iterate only once regardless of the number of tickers:

```VBA
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value <> Cells(i - 1, 1) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
            
        If Cells(i, 1).Value <> Cells(i + 1, 1) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
            
    Next i
```

After refactoring in this way, run times for 2017 data dropped to approximately 0.3 seconds and 2018 to approximately 0.4 seconds (on the same hardware).

## Summary

Refactoring code does not add new functionality, but can improve the ability to use or maintain existing code. In this case, by refactoring the code, our analysis is significantly faster and therefore more likely to scale as the amount of source data increases, i.e. to cover the entire stock market. Furthermore, since we separated the process of anaysis from presentation, it should also be easier to modify the output format if needed in the future.

Advantages of refactoring code might include better readability, better modularity (for code re-use), improved error handling, or better performance. It can also help you better understand the code you are trying to refactor. Such advantages, however, come at the cost of spending time modifying already working code. This process can introduce bugs or simply incur opportunity cost where that time might be better spent working on something new.
