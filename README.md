# stock-analysis
# Overview of Project

Steve, a recent finance graduate wanted to perform analysis on green energy stocks in order to help his parents diversify their investment ,since  they were only interested to invest in green energy stocks. Steve's parents invested only in DQ. To give other recommendations to his parents, we wrote a code for Steve to find two key parameters for green energy stocks  in 2017 and 2018 ,which were Total Daily Volume, that measures how frequently the stocks were traded and their Net Return.

In order to help Steve analyze and find these parameters for green energy stocks, a sample data set of 12 stocks containing various data in 2017 and 2018 was utilized. The Total Daily Volume and Net Return was calculated for each of the Green Energy stocks in the sample data set using VBA in Excel to automate and calculate the parameters accurately. 

Steve wants to use the same code to analyze the entire stock market which will have a lot more stocks than just the sample size of 12 that was initially analyzed. In order to make the run time for the code shorter to analyze the entire stock market, the initial code written was refactored. This report analyzes the green energy stock performance and the observations in run time of the new refactored code versus the old code for 2017 and 2018.

# Results
This is the observations of stock performance and run time in 2017 and 2018.
 
From our analysis on Green Stocks in 2017 , we can observe that
   * out of the 12 stocks analyzed , 11 of them had a postive return .(that is 91.6% success.) 
   * DQ was the most successful one with a return of 199.45% even though it had the lowest Daily Volume , that is was traded less frequently.
   * refactored code run time was reduced to 0.21sec from 1.14sec.
   
![VBA_Challenge_2017](https://github.com/maryamt95/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

From our analysis on Green Stocks in 2018, we can observe that
  * Out of the 12 stocks analyzed , only 2 had a positive return .(that is 16.6% success.)
  * DQ performed poorly  with a negative return of 62.6% .
  * RUN had the highest retun of 83.95% and was also traded most frequently.
  * Refactored code run time was reduced to 0.27sec from 1.09sec.
  
 ![VBA_Challenge_2018](https://github.com/maryamt95/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)
    
# Summary
From our analysis ,we can conclude that refactoring code did reduce the run time significantly in both 2017 and 2018. The run time reduced by 81% in 2017 and 75% in2018.
refactoring a code helps to improve the run time , and also help provide a better understanding of the code . 
In the VBA script , creating a tickerIndex for iterating all  the rows had made the code faster . this will help Steve greatly when he decides to use the code to analyze the entire stock market . The disadvantage being , it had made the code complex and can make it hard to understand.

