# Stock-Analysis

## Overview:
This project used VBA within Excel to analyze the performance of 12 stocks in 2017 and 2018.

### Purpose
The purpose of this project is to write VBA code (within an Excel file) that would, for each stock, calculate the starting and ending prices and then calculate the daily return for both 2017 and 2018.
As part of the project we were asked to refactor the initial code we created during the module exercise.  

## Results

### Data Analysis
In reviewing the returned data for both years, we're able ascertain that 2017 was a much better year for energy related companies in the stock market (if we assume that the performance of this set of stocks are representative of the overall market) than 2018.  In most cases, both the total daily volume and the return are down.
While most stocks daily volumes and return decreased in 2018, I found it interesting that the daily volumes for DQ nearly tripled in 2018 over 2017 yet the return dropped significantly. In comparison, the daily volumes for RUN nearly doubled in 2018 and the return significantly increased. 

![2017 Results](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/2017_Results.png)
![2018 Results](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/2018_Results.png)

#### How it was done
We used VBA code to identify each individual stock (ticker) and then using that information, loop through the rows to identify both the starting and ending values which were used to calculate the total daily volumes and return.
##### Original Code
The original code resulted used a nested for loop and if/then statements that went through each ticker, each row, returned the output for that row and then went to the next ticker.

 '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
    
    '5) loop through rows in the data
    Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then
        
                 totalVolume = totalVolume + Cells(j, 8).Value
         
                End If
           '5b) get starting price for current ticker
           
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
                StartingPrice = Cells(j, 6).Value
        
                End If
           
           '5c) get ending price for current ticker

              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
               EndingPrice = Cells(j, 6).Value
                
              End If

       Next j
       '6) Output data for current ticker

  Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = EndingPrice / StartingPrice - 1
    
    Next i

##### Refactored code
The refactored code used an index to eliminate the nested for loop, used one for loop and several if/then statements to run the function for each ticker than ran another for loop to return the outcome.

    '1a) Create a ticker Index
       Dim tickerIndex As Integer
       
        'Set ticker index to zero
        tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         'End If
         End If
            
                 
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
                        
        'End If
            End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)
    Next i


 
#### VBA Refactoring Impact
Refactoring this code resulted in about an 85% reduction in the amount of time it took for the code to run (as seen below).
|   Code     |       2017        |       2018        |
|------------|-------------------|-------------------|
|Original    | 1.089344 seconds  | 1.085938 seconds  |
|Refactored  | 0.1445313 seconds | 0.15625 seconds   |   

![2017 original run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/Original_code_2017_runtime.png)
![2018 original run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/2018_original_run_time.png)

![2017 refactored run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/2017RunTime.png)
![2018 refactored run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Resources/2018RunTime.png)


## Summary

### 1. Advantages / Disadvantages of refactoring code:
#### Advantages
- Refactoring code can reduce the runtime to perform the task
- Refactoring can shorten the overall length (number of lines) of the code
- Refactoring can make it easier for another person to step into the code (provided you have included appropriate comments)

#### Disadvantages
- If you refactor without keeping an understanding of the ultimate desired outcome, you could be too restrictive in the code.  You can't just refractor with the current step in mind.  You have to prepare the code to perform from beginning to end.
- Refactoring can make individual tasks more complex as single-action items.

 
### 2. How do these pros and cons apply to refactoring the original VBA script?
#### Pros
The advantages of refractoring, in this case are the reduce run time and shortening the length of the code.  If another person were to step into this code, they should be able to follow the intended function of the code.

#### Cons
The complexity of some of the lines of code may be a con, especially for new coders.
