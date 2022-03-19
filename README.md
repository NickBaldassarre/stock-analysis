# An Analysis of Stock Performance

## Overview of Project

### In order to help a friend advise his parents on which stocks to invest in, I will be using Excel Macros to organize and visualize the data in a way that will expose stock trends.

## Analysis

### Original Code Using Only One Array

In my first attempt at analyzing the data by year, I created a user input to choose the year to analyze, and an array ("tickers") containing the trading name of each stock. I created a variable ("ticker") to hold the value of the "tickers" array in a for loop, and then created a nested for loop that went over every row in the dataset to accumulate the total volume, starting price and ending price for each respective stock. These values were determined using if statements, and then outputed by the outer for loop. The final nested for loop is below.



    For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0

      For j = 2 To RowCount

      Worksheets(yearValue).Activate

        If Cells(j, 1).Value = ticker Then     
          totalVolume = totalVolume + Cells(j, 8).Value   
        End If

        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
          startingPrice = Cells(j, 6).Value
        End If

        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
          endingPrice = Cells(j, 6).Value
        End If

      Next j

    Worksheets("All Stocks Analysis").Activate
    Cells(i + 4, 1).Value = ticker
    Cells(i + 4, 2).Value = totalVolume
    Cells(i + 4, 3).Value = endingPrice / startingPrice - 1

    Next i

While this code was effective in analyzing the data, it took a very long time to do so. Each year took almost 20 seconds to analyze, pictured below.

![Original Runtime 2017](https://github.com/NickBaldassarre/stock-analysis/blob/5ce46350e96aeed9f4b09ce727e015c8de5ebbc7/Resources/Original_Code_2017.png)

![Original Runtime 2018](https://github.com/NickBaldassarre/stock-analysis/blob/5ce46350e96aeed9f4b09ce727e015c8de5ebbc7/Resources/Original_Code_2018.png)

### Refactored Code Using 3 Additional Output Arrays

In order to speed up the calculations taking place, I created 3 additional ouput arrays to hold the values as they were being determined. Now, instead of activating two different worksheets in every loop, outputing to one of them every time, the code just gathers the data from one worksheet into arrays. I am able to activate the worksheet containing the dataset before beginning the for loop, and can activate the ouput worksheets after the for loop. Refactored code below.

    For i = 0 To 11

      tickerVolumes(tickerIndex) = 0

      For j = 2 To RowCount

        If Cells(j, 1).Value = tickers(tickerIndex) Then
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
          tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        End If

      Next j

      tickerIndex = tickerIndex + 1

    Next i

After this, a simple for loop outputs all the values of each respective array. The result is impressive. What originally took almost 20 seconds has been reduced to a quarter of a second, pictured below.

![Refactored Runtime 2017](https://github.com/NickBaldassarre/stock-analysis/blob/bd0a47589b7f18d751544b25c9cf1267797de6fd/Resources/VBA_Challenge_2017.png)

![Refactored Runtime 2018](https://github.com/NickBaldassarre/stock-analysis/blob/bd0a47589b7f18d751544b25c9cf1267797de6fd/Resources/VBA_Challenge_2018.png)

### Results

After analyzing the data from both years, it is clear that only two stocks performed well: "ENPH" and "RUN". Both stocks are the only ones to see a return in 2018, each over 80%. When looking at 2017 however, "ENPH" performed far better, more than doubling in value, while "RUN" only increased by 5.5%. I would advise investing 75% into "ENPH" and 25% into "RUN" based on this dataset. That being said, I would recommend analyzing a far larger dataset before making any investment decisions.

### Summary

There are clearly advantages and disadvatages to refactoring code. One obvious advantage is being able to speed up processes that could otherwise take too long to be useful when dealing with much larger datasets. Another advantage could be when it is used as a first attempt at addressing code issues, before moving to bug fixing, which can take a lot longer. One major disadvantage could be that the refactored code may not have the same functionality. It can also be more difficult to refactor inefficient code than it would be to write it again from scratch.

In the VBA script for this project, the main advantage to refactoring the code is the amount of time saved in the calculations. It is truly amazing that such a seemingly small change in the code could have such a profound impact on performance. I did not find any obvious advantages to the original code used, especially since it took so much longer to run. This project certainly painted refactoring in a very positive light.
