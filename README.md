# Refactoring stock-analysis

## Overview of Project

### Purpose
The purpose of this project is to refactor vba code so that it will run faster, improve performance, and improve the user experience. Improved performance will be important as the dataset grows. To determine if there was an increase in performance, metrics will be taken for completion time.
### Results
After refactoring the code to run the stock analysis, there was a significant improvement in performance. Comparing running the old code with the running of the refactored code we can see a reduction of the time that it takes from execution to getting the results. The original code took about a second to run which doesn't seem bad.

![This is an image](/Resources/VBA_Challenge_2018.png)

The refactored code took less than a tenth of a second of a second to run.

![This is an image](/Resources/VBA_Challenge_2018_Refactored.png)

While the difference between 1 second and 1/10 of a second would not be noticeable for most, the difference between 10 minutes and one minute will make a noticeable difference.  This will be important when the data scales up.

This performance increase is due to the removal of an inner loop from the original code. The original code was written like this:
  ```vba
  '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
```
We can see that the first loop "i" loops 12 times and that the second loop, "j" loops to the count of all rows which is 3,012. Multiplying 12 x 3012, we get 36,144 loops.

The refactored code loop looks like this:
  ```vba
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            tickerIndex = tickerIndex + 1
            
        End If

    
    Next i
```
The refactored code loops through the 3,012 rows just once. This was done by using a TickerIndex variable to keep track of the index instead of using a loop. So, the refactored code should run about 12 times faster. When we divide 1 second by 0.0898, we get 11, which is close enough.
### Summary

-What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code are that it can save lots of time when running code which will lead to users have a much better experience using your code.  Log running code can drain resources and make other running programs slow as well as they compete for resources, refactoring code can minimize that risk.

Disadvantages of refactoring code is that the process can be very time consuming, and the results may not produce any performance increase, or the performance increase may not be of any significance.

-How do these pros and cons apply to refactoring the original VBA script?
In this case, we know that data will increase as time goes and more years are added. The time that was that was put into this project was minimal but had a significant impact on the time it took to run the code. The benefits outweighed the time that it took to refactor the code.

