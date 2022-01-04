# stock-analysis

## Purpose
The purpose of this project was to refactor code in Microsoft Excel VBA for stock data from 2017 and 2018. After analyzing the stock we were to find out if it was worth investing in any of the stocks in the analysis. 

The data we were given was from an earlier exercise, and the purpose of this challenge was to take that code and make it more efficient. After refactoring the code, the code was supposed to run more efficiently. 

We analyzed 12 different stocks to see which would be the most profitable for us to invest in. The information from the stocks contained the ticker, the date they were issued, the highest and lowest price, the volume of the stock, and the opening and closing prices. The purpose of the code we created was to retrieve the total daily volume, the ticker, and the return on each stock. 

### Results
Before refactoring the code, we were given a coding template to help us when refactoring the code. The module required us to fill in the rest of the code for parts: 1a, 1b, 2a, 2b, 3a, 3b, 3c, and 4.

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
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
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickersIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
                
                tickerIndex = tickerIndex + 1
            
        'End If
        End If
            
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
#### Summary

The advantages of refactoring code are pretty clear. They can help make your code much more efficient and easier to edit if you need to go back in and change anything. The increase in efficiency also allows you to read your code much easier. Most of the time, refactoring code will allow your code to run much faster.

The disadvantage to refactoring your code is it can cause errors. This is the problem I ended up running into and was unfortunately unable to figure this out. The error that I continued to run into were the following: 

![VBAErrorOutofRange](https://user-images.githubusercontent.com/95515322/148006952-19ba80f2-c9e2-49aa-b172-40681d7c9cb5.png)]

![VBAErrorHighlightedCode](https://user-images.githubusercontent.com/95515322/148007199-612298c6-6eb2-400c-81ab-57ea8ec054d0.png)]

