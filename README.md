# Analysis of Stock Data with VBA

## Overview of Project
An analysis of stock performance for 12 different Green Energy stocks using VBA.

### Purpose
The purpose of this project was to analyze the performance of 12 different Green energy companies to determine how one company (ticker name DAQO) compared to its peers. If DAQO performed poorly, the next task was to identify other companies from the target list for investment. To gauge company performance stock data from 2017 and 2018 was used. The final goal was to refactor the code to increase performance to potentially be used with larger datasets.

## Results
Initial analysis of DAQO performance revealed a loss of 62.6% in 2018, signaling it was not a good investment. Following this, I wrote an initial set of code to analyze the total volume of stocks traded and the percent return on those trades for 12 different Green Energy stocks in 2017 and 2018. Of the stocks analyzed, 'ENPH' and 'RUN' netted the highest consecutive returns, highlighting them as the best overall investment. This initial analysis took ~0.723s for 2017 and ~0.754s for 2018, "s" standing for seconds. Following this I refactored the code to make it more efficient, with the end result being ~6x faster. The process of refactoring is detailed in the below 'Analysis' section'

### Analysis
I began the refactoring process by creating a roadmap of the steps I would need to take for the code to function as intended. I then copied over the basic code that likely wouldn't need to be changed, being the headers, input box, ticker array and worksheet activation before placing them in their appropriate positions. This done, I reviewed the code again and determined the most optimal way to increase effeciency would be to reduce the number of nested for loops in the code. To do this, I created a new variable to hold the ticker array, and then created a series of arrays to hold ticker volume, starting price and ending price before setting the value of all these fields to 0. I then used these fields to write another for loop that would be able to read the data all at once before filling in the required fields, rather than having to loop through the data for each new stock price. 

The code used is as follows:
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    YearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrice(i) = 0
        tickerEndingPrice(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
        
            
        
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

End Sub


### Analysis of Outcomes Based on Goals

### Challenges and Difficulties Encountered

## Results
The net result was a nearly 6x increase in speed, with analysis time decreasing from 0.723s to 0.109s and 0.754s to 0.125s for 2017 and 2018 respectfully (as shown by the pictures below).

![VBA_Challenge_2017](https://github.com/Tbrecke01/VBA_Challenge/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/Tbrecke01/VBA_Challenge/blob/main/VBA_Challenge_2018.png)

###Pros and Cons of Refactoring
The biggest advantage for refactoring code is in its efficiency. The first and most obvious efficiency increase is through code run-rate, with succesfully refactored code running faster and (ideally) taking up less resources to complete. The other upside is efficiency in understanding. A cleaner code with less lines and (ideally) more explanation of steps makes it much more palatable for others to read, understand and use. This means that succesfully refactored code can more easily be taken and used as a model for projects outside of its initial function, whereas this may not be possible for long and messy code that has never been refactored.

The downside of refactoring is the cost of human time. While making the code simple, fluid and easy to digest may be ideal, it may not be feasible to do when facing strict deadlines. If the emphasis is placed on getting the code to work as fast as possible, irregardless of how it looks, refactoring could be a luxury a coder isn't able to enjoy. 

###Pros and Cons of Refactoring for this Project
The ~6x decrease in runtime for the code is a clear benefit. It is also easier to follow what the code is doing compared to the previous iteration.

The downside may be that the code has become too specialized for this particular task, and may be harder to utilize for other projects.
