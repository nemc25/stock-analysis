Stock-analysis
Module 2 VBA Challenge
Overview of Project: Explain the purpose of this analysis.
The purpose of this analysis is to determine if refactoring makes the VBA script runs fastar than the code done in green stocks xlsm file. This is done for the purpose of give to Steve a workbook where he can expand the dataset and include the entire stock market over the last few years. 

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Using the original script to run the stock performance between 2017 and 2018 I had to use more steps than in the refactored script. In the refactored script I used 4 different arrays: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. These were used as symbols to represents the stocks that were shown in the worksheet. tickerIndex was the name of the variable that connected the other three arrays. This aid to run faster the stocks. In the picture bellow, it shows how long it took to run the worksheet 2017.
![alt_text](https://github.com/nemc25/stock-analysis/blob/main/VBA_Challenge_2017.png)
Below it shoes how long it took to run the workseet 2018 with the new refactoring code:
![alt text](https://github.com/nemc25/stock-analysis/blob/main/VBA_Challenge_2018.png)



Summary: In a summary statement, address the following questions.

What are the advantages or disadvantages of refactoring code?
Refactoring script can makes the code run faster, this is an advantage of refactoring code. The disadvantage is that you have to add more code in VBA and it can make you more suceptible to error.

How do these pros and cons apply to refactoring the original VBA script?
The pros is that VBA original script is also used in refactoring. It is important to save the original script to avoid losing the code already made. If it is lost, it is extra work and may risk failure, and extra time.
