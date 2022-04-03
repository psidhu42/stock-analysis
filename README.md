# An Analysis of Stocks Using Excel VBA

## Overview of Project

Analyzing green energy stocks to determine their performance using Excel VBA.

### Purpose

The purpose of this was to refactor Excel VBA code, which looped through stocks data to determine if those stocks were worth investing in, and whether refactoring the code successfully made the VBA script run faster.

## Analysis Results

The original VBA code processed the data for 2017 and 2018 stocks at a much slower rate than the refactored code. The refactored code also included all the lines required for formatting the output so it was much more easily readable on the Excel sheet.

### Results of Original VBA Script

As shown in the following images, the original code looped through the stocks data and output the results onto the Excel sheet, notice that the data is not formatted in the cells. The times shown in the message box afterwards are also provided, with 2017 and 2018 results taking roughly 0.754 seconds each.

![2017 Results Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/module_2017.PNG) ![2018 Results Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/module_2018.PNG)
![2017 Runtime Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/runtime_2017.PNG) ![2018 Runtime Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/runtime_2018.PNG)

### Results of Refactored VBA Script

The refactored code performed much better than the original, with 2017 results taking around 0.109 seconds and 2018 results doing slightly better at about 0.105 seconds. The refactored code also includes formatting, and still did better it terms of time taken to output the results.

![2017 Results Refactored Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/refactored_2017.PNG) ![2018 Results Refactored Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/refactored_2018.PNG)
![2017 Runtime Refactored Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/VBA_Challenge_2017.png) ![2018 Runtime Refactored Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

## Summary

Refactoring code can have many advantages and disadvantages. Speeding up the time it takes to run the code is one advantage as shown in the results above. Some other advantages could be streamlining the code and make it easier to use with larger data sets or making it simpler to follow and understand for other and yourself. Disadvantages could be ones such as time constraint or time waste, for example if the code you are refactoring will only be used a few times, it would not be worth the trouble to edit it when you could just run the original and save many minutes or hours depending on the complexity of the refactoring steps needed.

In this case, refactoring the code was useful as it provided the advantage of speeding up the runtime by roughly 0.647 seconds, and when more stock data is added to the analysis that will significantly improve the time saved. The disadvantage of taking more time to refactor the code did arise for myself, as a new coder I found it confusing and difficult. It took me many hours of troubleshooting to finally get the refactored script working properly.