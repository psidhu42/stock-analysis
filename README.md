# An Analysis of Stocks Using Excel VBA

## Overview of Project

Analyzing green energy stocks to determine their performance using Excel VBA.

### Purpose

The puropose of this was to refactor Excel VBA code, which looped through stocks data to determine if those stocks were worth investing in, and whether refactoring the code successfully made the VBA script run faster.

## Analysis

The original VBA code processed the data for 2017 and 2018 stocks at a much slower rate than the refactorded code. The refactorded code also included all the lines required for formatting the output so it was much more easily readable on the Excel sheet.

### Results of Original VBA Code

As shown in the following images, the original code looped through the stocks data and output the results onto the Excel sheet, notice that the data is not formatted in the cells. The times shown in the message box afterwards are also provided.

![2017 Results Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/module_2017.PNG) ![2018 Results Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/module_2018.PNG)
![2017 Runtime Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/runtime_2017.PNG)      ![2018 Runtime Original Code](https://github.com/psidhu42/stock-analysis/blob/main/resources/runtime_2018.PNG)

### Analysis of Outcomes Based on Goals

When looking at the data for project outcomes based on the goal, the chart below shows that projects with a goal less than or equal to $5000 had a higher success rate than most projects above that goal.

![Outcomes vs Goals](https://github.com/psidhu42/kickstarter-analysis/blob/main/resources/Outcomes_vs_Goals.png)

## Limitations and Recomendations

The Kickstarter data set provides a lot of useful information about how well or not so well campaigns did in raising enough money to meet their goals. One of things it lacks is the ability to determine public interest in Louise's play specifically, this limitation could potentially lead to a failed campaign even if she reduced the budget and launched during a month that was otherwise full of successful projects.

Other useful tables and graphs could be ones that look at the number of backers per campaign or the average amount of donations to determine how many people Louise would potentially have to reach out to for a successful outcome. A chart that looks at the length of successful campaigns could also be useful in setting up an ideal deadline.