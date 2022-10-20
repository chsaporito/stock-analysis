# Refactoring stock-analysis

## Overview of Project

### Purpose
The purpose of this project is to refactor vba code so that it will run faster, improve performance and improve the user experience. Improved performance will be important as the dataset grows. To determine if there was an increase in performance, metrics will be taken for completion time.
### Results
After refacoring the code to run the stock analysis, there was a signifivcant improvement in performance. Comparing running the old code with the running of the refactored code we can see a reduction of the time that it takes from execution to getting the results. The orinal code took about a second to run which doesn't seem bad.

![This is an image](/Resources/VBA_Challenge_2018.png)

The refactored code took less than a tenth of a second of a second to run.

![This is an image](/Resources/VBA_Challenge_2018_Refactored.png)

While the difference between 1 second and 1/10 of a second would not be noticible for most,  the differnce beween 10 minutes and one minute will make a noticeable difference.  This will be important when the data scales up.

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
```vba


### Summary

-What are the advantages or disadvantages of refactoring code?

-How do these pros and cons apply to refactoring the original VBA script?
