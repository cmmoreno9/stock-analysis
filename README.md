# Written Analysis of Results

## **1) Overview of Project**
  The purpose of the analysis was to edit (refactor) the code used in the module in hopes of a faster execution time. In the future, Steve would like to include the entire stock market, and the original code would not be as effective and would take a longer time to evaluate. The refactored and original code evaluated 12 different tickers, total daily volume, and return rate dependent on the year. 

## **2) Results**

### Refactored Code 
![This is an image](https://github.com/cmmoreno9/stock-analysis/blob/43e166c52038a3ef5b50473e5f094238d2039829/Screen%20Shot%202022-04-02%20at%203.04.22%20PM.png)
 - In order to complete deliverable one, I created a tickerIndex variable and set it equal to zero. After, I created three output arrays: tickerVolumes, tickerStartingprice, and tickerEndingprice. TickerVolumes was set as a long data type, while the rest were set to single data types. There are 12 different tickers to evaluate; thus the output arrays are Dim (array) (12) As Single.  
 -Moving on to deliverable two, I started to create a for loop that initializes the tickerVolumes to 0. Additionally, I set values for the tickerIndex from 0-11 to account for the number of tickers. Although there are 12 tickers, the counting starts at 0. Thus the range will be from 0-to 11. 

![This is an image](https://github.com/cmmoreno9/stock-analysis/blob/43e166c52038a3ef5b50473e5f094238d2039829/Screen%20Shot%202022-04-02%20at%203.04.46%20PM.png)

 - Deliverable 3 is broken up into four sections. For the first section, I increased the tickerVolumes using the tickerIndex inside the for loop made in deliverable 2. Then I began to set my conditions. The first to check if the current row is the first with the selected tickerIndex, and another to check if the current row is the last row. If the next row doesn't match the ticker index, you increase the tickerIndex. After. its necessary to close the inner loop (I chose j), and then close the outer loop (tickerIndex). 
 -  In order to complete the last deliverable, I lopped through the 12 arrays and activated the output worksheet. I used the cells function to input the tickers, tickerVolumes, and the return value. 

![This is an image](https://github.com/cmmoreno9/stock-analysis/blob/43e166c52038a3ef5b50473e5f094238d2039829/Screen%20Shot%202022-04-02%20at%203.05.02%20PM.png)


### Original Run Times

![this is an image](https://github.com/cmmoreno9/stock-analysis/blob/d9cd74dad98933098e1b87abc5e927449ef5277b/Previous2017.png)
![this is an image](https://github.com/cmmoreno9/stock-analysis/blob/d9cd74dad98933098e1b87abc5e927449ef5277b/Previous2018.png)

### Refactored Run Times 
![This is an image](https://github.com/cmmoreno9/stock-analysis/blob/d9cd74dad98933098e1b87abc5e927449ef5277b/VBA_Challenge_2017.png)
![This is an image](https://github.com/cmmoreno9/stock-analysis/blob/d9cd74dad98933098e1b87abc5e927449ef5277b/VBA_Challenge_2018.png)

Overall, the goal was to reduce the run time to improve efficiency. The goal was met in comparing the previous and refactored run times as the refactored run times are significantly faster. 

## Summary 

1) One of the advantages of refactoring code is that it is easier to read, and it can also increase code efficiency. Additionally, if the code is easier to read and follows a more logical flow, errors are easier to pinpoint. Thus, maintenance isn't much of a hassle. However, refactoring code can also negatively affect the testing outcomes. Additionally, if it's a complex and lengthy code, refactoring can be incredibly time-consuming due to the constant functionality testing. 
2) Regarding how these drawbacks and advantages apply to this project, the refactored code reduced the runtime by decreasing the memory needed to process all the tickers at each row. This reduction was possible by optimizing the number of loops. This is the advantage. Although more due to human error, a disadvantage is the negative display of testing outcomes. I refactored my code, made an error in my loops, and did not receive the desired results. Finding the specific issue and making the correction was time-consuming. Thus, the possibility of negative test outcomes while in the process of refactoring is a drawback. 
