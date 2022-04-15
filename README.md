# Stock Analysis with VBA
Using VBA to create a macro in Excel that can be used to analyze stock market data. 

## Overview of Project

### Purpose
Using the macro created for this project, we can find and record the Total Daily Volume of stock sold for a particular stock symbol ticker, and also the Return of that stock over a year. This can be used to compare the performance of different stock purchases in order to make more informed decisions about stock that you should invest in. Over the course of the project, the code was refactored (cleaned and streamlined) to run as fast as possible so that the macro can be used on larger data sets as well. The macro can be used to analyze stock data from any year, and it can be run by simply clicking a button in your Excel Workbook.

## Results

### Analysis
In this project, the Excel Macro was used to analyze stock data from the years 2017 and 2018. We analyzed 12 different stocks tickers to find the total daily volume traded and the return on investment in each year. The tickers that had consistently postive return in 2017 and 2018 were ENPH and RUN. RUN had 5.5% return in 2017 and 84.0% return in 2018. If this growth continues into 2019, RUN would turn out to be alucrative investment.

![Stock_Analysis_2017](https://user-images.githubusercontent.com/100658772/163608565-40609444-73d3-4944-87ff-6612993a2946.png)

The analysis showed that in 2017, 11 out of 12 of these tickers had a positive return, with DQ having the highest return at 199.4%. Before deciding that DQ is an excellent investment, it is important to look at the data from 2018.

![Stock_Analysis_2018](https://user-images.githubusercontent.com/100658772/163608842-9b9233bb-c201-4c21-8de0-1e651bd1af90.png)

In 2018, DQ's return plummeted to -62.6%, which is the greatest loss compared to the other tickers analyzed. ENPH and RUN grew significantly from 2017 to 2018. More information about each industry would be valuable in determining why DQ's return dropped so low in 2018.

### Code Review
As part of this project, we refactored code to run faster and more efficiently so that larger data sets can also be analyzed easily. Refactoring the code had the impact of cutting down the run time for analyzing 2017 stock data from .9648438 seconds to .1601563 seconds and analyzing the 2018 stock data from .9882813 seconds to .15625 seonds. 

2017 Before Refactoring:
![2017_Before_Refactor](https://user-images.githubusercontent.com/100658772/163621751-6842e8c3-ad93-4768-9320-b067db98c62c.png)

2017 After Refactoring:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/100658772/163621777-5b8c9896-ef76-4723-a986-4ee8c72b45a4.png)

2018 Before Refactoring:
![2018_Before_Refactor](https://user-images.githubusercontent.com/100658772/163621804-a5f9f317-26ea-40a8-bf1f-9aa8040ad983.png)

2018 After Refactoring:
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100658772/163621815-47680d17-c7cc-4d06-87f9-fa9f6a9d78b1.png)

The refactored code, with comments, can be seen here:
[AllStockAnalysisRefactored.pdf](https://github.com/Giusi007/stock-analysis/files/8497786/AllStockAnalysisRefactored.pdf)

### Summary
The advantages of refactoring code is that it can make the code run more efficiently, and also be more understandable to yourself or other who are working on it in the future. A disadvantage could be that refactoring code takes additional time, but focusing on efficiency and clarity when you are writing code saves a lot of time and errors in the future if you are using the code over and over again.

For example, in this project, refactoring the VBA script impacted the run time of the script. Although cutting time down by les than one second may seem trivial, the refactored code would have a greater impact on large data sets. Creating multiple different arrays and variables and coding with them could result in some confusion if someone else is looking at the script or working on it in the future, so it is important to add comments explaining what each section of code is meant to be doing.
