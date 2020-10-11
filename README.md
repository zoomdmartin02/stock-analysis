# Stock-analysis Excel Application and Code Refactoring Analysis
---
The Deliverable Fulfills "Emerging" Required Criteria AND has the following: ✓There is a detailed summary of the pros and cons of refactoring code. ✓There is a detailed summary of the pros and cons of the original and refactored VBA script.


## Project Overview
---
This project takes two versions of an Excel Stock Analysis workbook and compares the VBA code to determine which is most efficient.
---
### Purpose
---
In working with data sets, size matters.  Code utilized to process data sets has to be efficient to create the most utility for the end-user so that time is not wasted in waiting for results to be presented.  Further, computing resources with inefficient code could be in danger of being over-leveraged, creating stability issues in the compute environment.  This project looks at two versions of VBA code, an original version and a refactored version to see if there are any code techniques that produces a more efficient experience.
---
### Background
---
The original and refactored code versions of this Excel worksheet produce the exact same results, which is a formatted worksheet with analysis of 12 different stock tickers.  Each of the stock tickers in the data set are evaluated and volumes tabulated and rates of return provided for each ticker symbol.  In each version, the user is prompted to select with data set to process, 2017 data or 2018 data.  Conditional formatting shows positive returns in cells with a green background while negative returns have cell backgrounds in red.
---
### Technology Involved
---
The technology involved if VBA code, the macro language for Microsoft Office.  In this workbook, there are no non-programmatic activities.  All processing and programmatic execution is occurring within VBA modules that are accessed thorugh the Developer ribbon within the application.
---
#### Original Code
---
The original VBA code relies on nested For Loops and conditional compound If Then and Else statements.  The outer loop cycles through a single array of 12 ticker symbols.  ?The conditional statements are utilized to establish a comparison between the ticker symbol in the data set and the current ticker in the variable array.  If the ticker symbol is the same, in an additive manner an update to variables occur as the program loops through from the inner loop each line of data in the data worksheet.  As it progresses through the data set a conditional statement determines whether the stock ticker changes.  If so, the values of the variables and calculations that determine return, starting and ending prices are writen to the analysis worksheet, and then processing returns to the code to progresses to the next ticker.  This process continues through the entire data set pausing each time to write summary data to the Excel worksheet after reading the last line item associated with the current stock ticker. Following is a screen shot of the key code snippets:
---

---
#### Refactored Code
---
The second version of this VBA code takes a different strategy with processing the data set.  It has all of the same techniques for collecting user input to select the appropriate data set and sheet formatting, but instead of collecting results into a single set of variables that are populated and written to the sheet with each subset of ticker symbols, this code creates variable arrays and an index variable array so that instead writing data to the sheet multiple times, data is written to the arrays in memory and only one write to the excel spreadsheet after the entire data set is processed and all elements of the arrays are populated and calculated.
---
## Results
---
![Table-1 ](Resources/Table_of_Outcomes_by_Goal.png)
## Summary
