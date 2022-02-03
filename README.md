# Refactoring VBA Code

## Overview of Project

### Purpose
The purpose of this project was to refactor VBA code that analyzed different stocks over the course of a year and make it more efficient. 
## Results

### Comparing original and refactored VBA Code
To compare the orginial and refactored VBA code, we will start by examining the similarites between the two. The codes are the same in the beginning. They both create start and end time variables to track the time it takes to execute the whole code. They also both create and initialize an array of all the tickers. Lastly, both codes format the analysis table the same.
Now, we can move on to the differences between the two codes. The differences in the two codes are mainly in the body of the codes. The original code uses an outer for loop to loop through all the tickers. Inside this for loop is another for loop that loops through all the rows in the data. The refactored code uses one for loop at the beginning to initalize the ticker volumes to zero. Then it uses a second for loop (not inside the last) to loop through all the rows in the spreadsheet. When doing this it uses the ticker index to check for the last row with the selected ticker. When the last row is found, it increases the ticker index so it can move on to finding other values it needs to store for our analysis table. The refactored code then has a third for loop (not inside the other two) that places the values in our analysis worksheet.
The purpose of refactoring the code to have multiple separate small for loops instead of one large for loop within another is because the code will only have to run through the necessary number of excel rows and stop when it finds the data it needs for each ticker. The original code was running through 3,000 rows 12 times, while the refactored code is just running through 3,000 rows once. The orginial code, refactored code, as well as execution times for each year and code can be seen below. 
#### Original Code
![original-code-image](https://github.com/nicole-peltier/stock-analysis/blob/main/original_1.png)
![origina-code-2](https://github.com/nicole-peltier/stock-analysis/blob/main/original_2.png)

#### Refactored Code
![refactored-code-1](https://github.com/nicole-peltier/stock-analysis/blob/main/refactored_1.png)
![refactored-code-2](https://github.com/nicole-peltier/stock-analysis/blob/main/refactored_2.png)

#### Original Code 2017 Execution Time
![original-code-2017](https://github.com/nicole-peltier/stock-analysis/blob/main/2017-original.png)

#### Original Code 2018 Execution Time
![original-code-2018](https://github.com/nicole-peltier/stock-analysis/blob/main/2018-original.png)

#### Refactored Code 2017 Execution Time
![refactored-2017](https://github.com/nicole-peltier/stock-analysis/blob/main/2017-refactored.png)

#### Refactored Code 2018 Execution Time
![refactored-2018](https://github.com/nicole-peltier/stock-analysis/blob/main/2018-refactored.png)

### Comparing Stock Performance between 2017 and 2018
After writing the necessary code, one can now quickly analyze the results of the stocks in 2017 and 2018. The written code executes a table including three columns: the ticker, total daily volume, and the return. For 2017 the results were as followed:
![stocks-2017-image](https://github.com/nicole-peltier/stock-analysis/blob/main/2017%20stocks.png)

Looking at this table, one can see all of the stocks except TERP have a positive return rate. For 2018 the results were as followed:
![stocks-2018](https://github.com/nicole-peltier/stock-analysis/blob/main/2018%20stocks.png)

Comparitively, a majority of stocks in 2018 saw a drastic decrease in return rates. Although ENPH still has a positive return rate, this stock when down 47.6%. However, two stocks saw in increase in return rates. These were RUN and TERP.
## Summary

- What are the advantages and disadvantages of refactoring code?
The advantages to refactoring code is making it more efficient. By doing this you are using less memory, taking fewer steps (which leads to faster execution times), and improving the logic of the code to make it easier for future users to read. The disadvantages of refactoring code is that it can be time consuming. Figuring out what someone elses ideas were and how to make it more efficient while still accomplishin the same result can be difficult.
- How do these pros and cons apply to refactoring the original VBA script?
When the original VBA script was refactored the execution time went down and the code was easier to follow logically. However, trying to figure out how to refactor this code took some time. The process to figure out how to index through the rows took many trial and errors, and a whole lot of debugging. But the pros outweighed the cons due to the ending result of the refactored code. 
