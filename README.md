# Stock-Analysis
# Stock-Analysis
## Overview:
This project used VBA within Excel to analyze the performance of 12 stocks in 2017 and 2018.

## Purpose
The purpose of this project is to write VBA code (within an Excel file) that would, for each stock, calculate the starting and ending prices and then calculate the daily return for both 2017 and 2018.
As part of the project we were asked to refactor the initial code we created during the module exercise.  

# Results
'Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script
## Data Analysis
In reviewing the returned data for both years, we're able ascertain that 2017 was a much better year for energy related companies in the stock market (if we assume that the performance of this set of stocks are representative of the overall market) than 2018.  In most cases, both the total daily volume and the return are down.
While most stocks daily volumes and return decreased in 2018, I found it interesting that the daily volumes for DQ nearly tripled in 2018 over 2017 yet the return dropped significantly. In comparison, the daily volumes for RUN nearly doubled in 2018 and the return significantly increased. 

<img src = "https://github.com/LauraZJ/Stock-Analysis/blob/main/2017_results.png" alt = "2017 Results" width = "300"/>
<img src = "https://github.com/LauraZJ/Stock-Analysis/blob/main/2018_results.png" alt = "2018 Results" width = "300"/>

## How it was done
We used VBA code to identify each individual stock (ticker) and then using that information, loop through the rows to identify both the starting and ending values which were used to calculate the return.

### VBA Refactoring Impact
Refactoring this code resulted in about an 85% reduction in the amount of time it took for the code to run (as seen below).
|   Code     |       2017        |       2018        |
|------------|-------------------|-------------------|
|Original    | 1.089844 seconds | 1.09375 seconds |
|Refactored  | 0.148438 seconds  | 0.15625 seconds   |   

![2017 original run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Original_code_2017_runtime.png)
![2017 refactored run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/2017RunTime.png)

![2018 original run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/Origina_code_2018_runtime.png)

![2018 refactored run time](https://github.com/LauraZJ/Stock-Analysis/blob/main/2018RunTime.png)

## Code comparison

# Summary


## Advantages / Disadvantages of refactoring code:
### Advantages
- Refactoring code can reduce the runtime to perform the task
- Refactoring can shorten the overall length (number of lines) of the code
- Refactoring can make it easier for another person to step into the code (provided you have included appropriate comments)

### Disadvantages
- If you refactor without keeping an understanding of the ultimate desired outcome, you could be too restrictive in the code.  You can't just refractor with the current step in mind.  You have to prepare the code to perform from beginning to end.
- Refactoring can make individual tasks more complex as single-action items.

 
## 2. How do these pros and cons apply to refactoring the original VBA script?

