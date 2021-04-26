# Stock_Analysis

## Overview of Project
The purpose of the project is to refactor previous code of stocks analysis in order to make it work more efficient for larger dataset.

## Results
After refactoring the code execution showing better efficiency. Following screenshots confirming that execution time significantly reduced: [VBA_Challenge_2017.png]P{https://github.com/andgerashchenko/andrey_stock-analysis/blob/3e41582b10fbae8e35178afc8c81874d437c0754/resourses/VBA_Challenge_2017.png} & [VBA_Challenge_2018.png]{https://github.com/andgerashchenko/andrey_stock-analysis/blob/3e41582b10fbae8e35178afc8c81874d437c0754/resourses/VBA_Challenge_2018.png}. The main idea of refactoring the code is to get rid of repitdely looping through whole set of data numeroous times, in other words to repace nested loops with other code, using Index function. In that case variable 'tickerIndex' was created and used in the script. For example when increase of total volume of stocks needed, instead of 'For' loop through all data, we run through it just once with the following code: 
'If Cells(i, 1).Value = tickers(tickerIndex) Then
 tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)'
 
##Summary
Code refactoring make it looks neat, easy to read and maintain. On the other hand it takes time, which can be crucial in some cases. Also refactoring can make the code more specific and less flexible.
In this particular script refactiring made the code work faster but if we will need to adjust it for changed data, the larger amount of code will be needed to rewrite.
