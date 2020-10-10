# Stock-analysis Excel Application and Code Refactoring Analysis
---
The purpose and background are well defined. ✓The analysis is well described with screenshots and code.The Deliverable Fulfills "Emerging" Required Criteria AND has the following: ✓There is a detailed summary of the pros and cons of refactoring code. ✓There is a detailed summary of the pros and cons of the original and refactored VBA script.


## Project Overview
---
This project takes two versions of an Excel Stock Analysis workbook and compares the VBA code to determine which is most efficient.
---
### Purpose
---
In working with data sets, size matters.  Code utilized to process data sets has to be efficient to create the most utility for the end-user so that time is not wasted in waiting for results to be presented, or that computing resources are not over-leveraged, creating stability issues in the compute environment.  This project looks at two versions of VBA code, an original version and a refactored version to see if there are any code techniques that produces a more efficient experience.
---
### Background
---
The original and refactored code versions of this Excel worksheet produce the exact same results, which is a formatted worksheet with analysis results where each of the stock tickers in the data set are evaluated and volumes tabulated and sticker rates of return provided.  In each version, conditional formatting shows positive returns in cells with a green background while negative returns have cell backgrounds in red.
---
### Technology Involved
---
#### Original Code
---
The original VBA code relies user input to identify the data set (e.g. which year, 2017 or 2018) is appropriate to process.  Nested For loops and considtional compound If Then and Else statements are utilized to establish the ticker symbol and if the ticker symbol is the same, in an additive manner update a variable as it progresses through the data set and then when it is determined that the stock ticker changes, the values of the variables and calculations to determine return, starting and ending prices are writen to the analysis worksheet and then processing progresses to the next ticker.  This process continues through the entire data set pausing each time to write summary data to the Excel worksheet after reading the last line item associated with the current stock ticker. 
---
#### Refactored Code
---
The second version of this VBA code takes a different strategy with processing the data set.  It has all of the same techniques for collecting user input to select the appropriate data set and sheet formatting, but instead of collecting results into a single set of variables that are populated and written to the sheet with each subset of ticker symbols, this code creates variable arrays and an index variable array so that instead writing data to the sheet multiple times, data is written to the arrays in memory and only one write to the excel spreadsheet after the entire data set is processed and all elements of the arrays are populated and calculated.
---
## Results
---
![Table-1 ](Resources/Table_of_Outcomes_by_Goal.png)
## Summary
