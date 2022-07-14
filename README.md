# Stock-Analysis

Analyzing Green Energy Stocks in VBA based on Financial Data

## Overview of Project

Steve’s parents are interested in investing in Green Energy Stocks and Steve needs our help in analysing various stock options. We analyzed DAQO and 11 other stocks with its daily volume and percentage change in prices to determine the performance of each stock. Now, Steve wants to expand the dataset to include the entire stock market to find better investment options for his parents.

With the current code, it is possible to analyze and get outcomes for few stocks. However, it may be difficult to analyze thousands of stocks with its daily information on volume, starting price, ending price and percentage change in prices. This can become complicated as well as time-consuming. 

To overcome this hurdle, we will be refactoring the code to make it more efficient. We will be initializing all variables and then looping the data once to get the same results we got from our original code. At the end, we will be comparing run time of both our codes to determine if the refactored code runs faster and thereby help Steve and his parents in finding a stock for investment.

## Original Code and its Timelapse

In our original code, first we initialised our tickers and looped the data through each ticker. Then, we created another loop within the main loop (nested loops). This loop will go through each row in dataset and check if the row is related to the specified ticker and if it is the desired ticker, then totalVolume variable will increase and we can get the startingPrice, endingPrice and percentage change in price to determine performance of that stock.  Below is a screen shot of the original code.
 

![Original_Code_Screenshot](https://user-images.githubusercontent.com/108366412/178903122-9732a525-fa27-450c-9411-6ca65ec46178.png)


The results shows that the total trade volume of majority of stocks has increased from 2017 to 2018 except stocks like AY, FSLR JKS and SPWR which had minor reductions in its volumes. Also, the performance of majority of stocks declined in 2018 with each stock registering negative % change except ENPH (81.9% increase) and RUN (84% increase). Hence, we can say that overall, the Green Energy Stocks are not stable in 2018 as although the volumes increased, the prices fell in 2018 as compared to 2017. Conditional formatting applied on the results helps us in reading these results in a glance.

The code ran successfully with the given data. However, due to the nested loop, the code took more time as it had to first check if it is the correct ticker and then fulfill the conditions going through the dataset repeatedly. Below are the screen shots of the run time it took for the original code for each of the years –
 
![Original_Timelapse_2017](https://user-images.githubusercontent.com/108366412/178903184-a0fafe69-f7e3-4206-9b5c-06cfb1d99590.png)

![Original_Timelapse_2018](https://user-images.githubusercontent.com/108366412/178903206-e4d21372-d752-45ce-9fb2-bb13e2d3c61a.png)

## Refactored Code and its performance:

Our main purpose for refactoring the original code is to make it easy to follow and run it faster. Applying our VBA knowledge in the Starter code provided will help us accomplish that goal. The dataset, variables and the required outcomes will be the same as original code.

Firstly, we initialized tickers and set a tickerIndex. Then, we initialized tickerVolume, tickerStartingPrice and tickerEndingPrice variables and formed a loop to run through the dataset to collect values for our variables related to that ticker. In this code, the loop will run just once hence it is expected to take shorter time than the original code. Below is the screen shot of the refactored code.


![Refactored_Code_Screenshot](https://user-images.githubusercontent.com/108366412/178903255-8266ec0f-32b8-4ec8-9087-eb05032548cd.png)


From the below refactored code screenshots it gets clear that it does take less time to run the updated code
 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/108366412/178903276-ce4b6570-344f-494a-9486-285a3189ee7c.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/108366412/178903282-ebeb9735-4e75-4a2e-8d7b-29b92d895060.png)


## Advantages of Refactoring

  * Code can be made less complicated with fewer conditions and loops.

  * Code runs much faster when refactored as unnecessary steps are removed. Time factor becomes important when working with large datasets. 
  
  * Code is easy to follow when refactored properly. 

## Advantage pertaining to the VBA script

With the refactored code, we were able to run the code within 0.4 seconds on an average for each year as compared to 2 seconds on an average for original code. If the dataset was larger, it would have made a huge difference. Also the code looks clean with just one loop running.  

## Disadvantages of Refactoring

  * Outcome of the analysis is compromised if vital steps are removed during refactoring.
  
  * The whole pseudocode needs to be altered and in doing so, code can get more complex and difficult to understand.
  
  * Sometimes it is not worth the time to refactor the code with small datasets.

## Disadvantage pertaining to VBA script

Although in the refactored VBA script, the loop runs once and it takes shorter timeframe to get the outputs, the code gets difficult to write and understand with the initialization of tickerIndex and the usage of tickerIndex variable in each of the if-then condition statements to find the outcomes.
