# Stock-Analysis - Module 2 Exercise
## Overview – VBA Challenge of All Stocks Analysis data set 
### Goal
### We are being asked to Refactor the green-stocks.xlsm data set to ensure it can collect the entire data set with one loop. The next step was to determine if the refactoring actually runs faster. Also, the goal was to ensure the script efficient and required less steps.
### Below are the steps the Module 2 Challenge required:
#### - Download the challenge_starter_code.vbs file and rename it VBA_Challenge.vbs.
#### - Create a folder called “Resources” to hold the run-time pop-up messages that you’ll screenshot after running refactored analyses for 2017 and 2018.
#### - Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm.
#### - Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
#### - Use the steps below to add code were indicated by the numbered comments in the starter code file.
### The Challenge -
#### Step 1a:
#### Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
##### Below are the steps taken to achieve the requested challenge. Each step is followed by the piec eof code that was used to solve the step in the challenge. We introduced a loop to esnure the enire data set is taken into account and then applied the requested logic.
##### Solution (1a) - Create a ticker Index
######    tickerIndex = 0
#### Step 1b: - Create three output arrays
#### Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickerVolumes array should be a Long data type. The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
##### Solution (1b)
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
#### Step 2a - Create a for loop to initialize the tickerVolumes to zero. If the next row’s ticker doesn’t match, increase the tickerIndex.
#### Create a for loop to initialize the tickerVolumes to zero.
##### Solution (2a)
       For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
       Next i
#### Step 2b - Loop over all the rows in the spreadsheet.
#### Create a "For Loop" that will loop over all the rows in the spreadsheet.
##### Solution (2b) - Loop over all the rows in the spreadsheet.
       For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
       Next i
#### Step 3a
#### Write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
##### Solution (3a) - Increase volume for current ticker
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
### Step 3b
Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.
##### Solution (3b) - Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
### Step 3c 
Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
##### Solution (3c) - check if the current row is the last row with the selected ticker.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
### Step 3d 
Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
##### Solution (3d) - Increase the tickerIndex
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
          Next i
### Step 4
Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
    For i = 0 To 11
      Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
#### Finally, run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module (as shown in the images below). In your Resources folder, save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook.   
##### ![VBA_Challenge_2017](https://user-images.githubusercontent.com/78942339/125214091-3bfd9600-e283-11eb-8186-4a800dd0299e.PNG)
##### ![VBA_Challenge_2018](https://user-images.githubusercontent.com/78942339/125214092-3ef88680-e283-11eb-889b-a0faf97def66.PNG)
### Summary
#### Refactoring Code in general
##### Has advantages and can provide simplified support and code updates. Clean code is much easier to update and improve. Refactoring also allows for saved time and money in the future. Code refactoring reduces the likelihood of errors in the future and simplifies the implementation of new software functionality. One of the major disadvantages of refactoring is it is very time consuming.
#### Refactoring the VBA script
##### Refactoring the code allowed us to look for more efficient code to run the loop over the entire data set taking the least amount of time possible. The new For loop is efficient and is minimized the most reduced steps available in VBA. One disadvantage was I found it difficult to start over, but in other areas we were able to reuse older code.                    
