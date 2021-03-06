# Stock-analysis Excel Application and Code Refactoring Analysis
---
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
The original and refactored code versions of this Excel worksheet produce the exact same results, which is a formatted worksheet with analysis of 12 different stock tickers.  Each of the stock tickers in the data set are evaluated and volumes tabulated and rates of return provided for each ticker symbol.  In each version, the user is prompted to select which data set to process, 2017 data or 2018 data.  Conditional formatting shows positive returns in cells with a green background while negative returns have cell backgrounds in red.
---
### Technology Involved
---
The technology involved is VBA code, the macro language for Microsoft Office.  In this workbook, there are no formulas or functions established in the worksheet itself.  All processing and programmatic execution is occurring within VBA modules that are accessed thorugh the Developer ribbon within the application.  Formatting, formulas and functions are called from the code and applied to cells or ranges from the VBA code.
---
#### Original Code
---
The original VBA code relies on nested For Loops and conditional compound If Then and Else statements.  The outer loop cycles through a single array of 12 ticker symbols.  The conditional statements are utilized to establish a comparison between the ticker symbol in the data set and the current ticker in the variable array.  If the ticker symbol is the same, in an additive manner an update to the variables occur as the inner loop code evaluates each line of data in the worksheet.  As it progresses through the data set, a conditional statement determines whether the stock ticker changes.  If so, the values of the variables and calculations that determine return, starting and ending prices are writen to the analysis worksheet, and then processing returns to progress to the next ticker whithin the loop.  This process continues through the entire data set pausing each time to write summary data to the Excel worksheet, after reading the last line item associated with the current stock ticker. Following is a screen shot of the key code snippets:
---
This is the tickers array from the original code.
![Figure-1 ](Resources/Original_Array.png)
---
This is the original loops and conditional statements.  Notice how it writes to the cells in the analysis sheet each time it loops through.
![Figure-2](Resources/original_loop.png)
---
#### Refactored Code
---
The second version of this VBA code takes a different strategy with processing the data set.  It has all of the same techniques for collecting user input to select the appropriate data set and sheet formatting, but instead of collecting results into a single set of variables that are populated and written to the sheet each revolution of the loop, this code creates 3 variable arrays and an index variable array so that instead of writing data to the sheet multiple times, data is written to the arrays in memory and only one write to the excel spreadsheet occurs after the entire data set is processed and all elements of the arrays are populated and calculated.  The following code snipets demonstrate this change in strategy:
---
This picture shows the additional arrays utilized in the refactored code.
![Figure-3](Resources/refactored_multiple_arrays.png)
---
Here again, looping through, but writing instead to the variable arrays using an Index array to keep track of each of the array positions.
![Figure-4](Resources/refactored_loops.png)
---
Now, we're looping through each of the arrays to write their values to the worksheet.
![Figure-5](Resources/refactored_looping%20_multiple_arrays.png)
---
## Results
---
Here are the screenshots that show the amount of time it took to run each of the two data sets:
---
2017 ![Figure-6](Resources/green_stocks_analysis_2017.png)
---
2018 ![Figure-7](Resources/green_stocks_analysis_2018.png)
---
Nearly 12.5 seconds to run the original code.  In total there are about 3000 lines of data in these subsets.  These data sets are also very small in comparison to an entire stock index's data such as the DOW Jones.  With 12 tickers, we took nearly 1 second per ticker to process.  Think about what this would look like if there were hundreds or thousands of stock tickers to be processed.
---
Here are the screenshots that show the amount of time it took to run each of the two data sets with the refactored code:
---
2017 ![Figure-6](Resources/VBA_Challenge_2017.png)
---
2018 ![Figure-7](Resources/VBA_Challenge_2018.png)
---
With the refactored code, we seem to have brought our processing time down significantly from nearly 12.5 seconds to now about half a second.  I had heard that processing in memory is always faster than writing to the hard drive, but this brings that concept to life.
---
## Summary
---
I can only think of one reason why a developer would not be interested in learning ways to improve his/her code.  Refactoring simply looks at possible improvements.  In our case above, we have shown that signifcant improvement to the user experience and resoure utilization can be achieved by refactoring code instead of always going back to original development efforts.
---
The original vs refactored code in this case showed significant improvement by just minimizing writes to the worksheet at the end and leveraging variable indexes, allowing all processing to stay in memory as opposed to writing each time it cycled through the outer loop.  Any client or customer would very likely greatly prefer the refactored over the original code, especially as data sets being processed increase.
---
