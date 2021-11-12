# `**Overview of Project: Explain the purpose of this analysis.**`
In this challenge,   I edit, or refactor, tsolution code to loop through all the data one time in order to collect the same information of stocks in 2017 and 2018 years. Then, I’ll determine whether refactoring my code successfully made the VBA script run faster. Finally, I’ll present a written analysis that explains my findings
# **Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.**
Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.

'1a) Create a ticker Index
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.
For j = 2 To RowCount
queryticker = Cells(j, 1).Value

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i


![All Stocks 2017](https://github.com/olenarabani/stock-analysis/blob/main/Code%20ran%202017.png)

![All Stocks 2018](https://github.com/olenarabani/stock-analysis/blob/main/Code%20ran%202018.png)

# **Summary: In a summary statement, address the following questions.**

## What are the advantages or disadvantages of refactoring code?
Disadvantages:
Refactoring costs development time. And may not be safe.
You may underestimate the amount of time for refactoring and end up working on it longer than you planned.
 it can takes time, and if done wrong, it you can create unnecessary tight couplings between unrelated modules of the system and makes things even more complex. If you don't have or create a good set of tests to back you up then you can break things too.
Advantages:
Refactoring facilitates easy/safe change in the future.
Understandable code when you come back to look or refactor further in a few weeks/months, plus extensible code.
You have a cleaner, extensible, easier to maintain code base. It will definitely lower the technical debt 

## How do these pros and cons apply to refactoring the original VBA script?
After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. Disadvantages of Code Refactoring: Time Consuming: You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go.
