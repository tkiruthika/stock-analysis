# Stock Analysis
## Overview
Refactoring the VBA script to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
### Purpose
The purpose of refactoring this code is to make the VBA script run faster. 
## All Stock Analysis
The stock data for the year 2017 and 2018 are given. We have to analyze the stock data to find
  - ### Total daily volume
     Daily volume is the total number of shares traded throughout the day. It measures how actively a stock is traded.
  - ### Yearly return for each stock
     The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
### VBA Script
- Macro **AllStockAnalysis** is created and **startTime** and **endTime** Variables are declared with the datatype **Single**.
- The output sheet **"All Stocks Analysis** is formatted as per the requirement.
- The Input box is created to get the input **yearValue** for which the analysis has to be done.
- The Header row for the out sheet is created using **Cells()** function.
- An array for all the **tickers** are initialized.
- Variables **startingPrice** and **endingPrice** are declared with the datatype **double**.
- Number of rows in the datasheet is calculated using **Rows.count**, to loop through the rows in the data.
- A **for** loop is created to loop through the **tickers** array and **totalVolume** is intialized to zero.
- A **nested for** loop is created to loop through the rows in the data.
- The **totalVolume**, **startingPrice**, and **endingPrice** are calculated after checking for conditions with the corresponding **if** statement
- The Outputs are displayed on the corresponding cells as required in the output sheet.
- Formatting of the output data is done as per the requirement by using different attributes like **FontStyle, Borders, NumberFormat, Interior.Color**, etc.


Sub AllStockAnalysis()
    
    Dim startTime As Single
    Dim endTime  As Single
    
    
    '1)Format the output sheet on the "All Stocks Analysis" worksheet.

    Worksheets("All Stocks Analysis").Activate
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row.
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2)Initialize an array of all tickers.
    
    Dim tickers(11) As String
    
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
    
    '3a)Initialize variables for the starting price and ending price.
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b)Activate the dataworksheet.
    
    Sheets(yearValue).Activate
    
    '3c)find the number of rows to loop over.
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4)Loop through the tickers.
    For i = 0 To 11

       ticker = tickers(i)
       totalVolume = 0
     
     '5)Loop through rows in the data
     
      Sheets(yearValue).Activate
       For j = 2 To RowCount
       
       '5a)Find the total volume for the current ticker.
       
       If Cells(j, 1).Value = ticker Then
       
            totalVolume = totalVolume + Cells(j, 8).Value
            
          End If
          
       '5b)Find the starting price for the current ticker.
       
       If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then       
            startingPrice = Cells(j, 6).Value          
          End If
       '5c)Find the ending price for the current ticker.
       
       If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
       
            endingPrice = Cells(j, 6).Value
            
          End If
       Next j
      '6)Output the data for the current ticker.
       
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i 
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    'Range("C4:C15").Style = "Currency"
    Columns("B").AutoFit   
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
      
      If Cells(i, 3) > 0 Then
            
            'Color the cell green
           
           Cells(i, 3).Interior.Color = vbGreen
       
       ElseIf Cells(i, 3) < 0 Then
            
            'Color the cell red
            
            Cells(i, 3).Interior.Color = vbRed
        Else
   
            'Clear the cell color
   
            Cells(i, 3).Interior.Color = xlNone
      End If
  
    Next i
    
      endTime = Timer
  
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)   

End Sub    
  
### Refactored VBA Script
The above VBA script is refactored by
- Creating **tickerIndex_** variable with the data type **Integer** and initialized to zero.
- Creating three output arrays **tickerVolumes** with the data type **long**, **tickerStartingPrices and tockerEndingPrices** with the data type **Single**.
- The **tickerVolumes**, **tickerStartingPrices**, and **tickerEndingPrices** are calculated using **tickerIndex** as the index of the arrays.



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
    Dim tickerIndex As Integer
    tickerIndex = 0   'tickerIndex is set to zero before iterarting over all the rows
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long   'data type for tickerVolumes is Long
    Dim tickerStartingPrices(12) As Single   'data type for tickerStartingPrices is Single
    Dim tickerEndingPrices(12) As Single     'data type for tickersEndingPrice is Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 12
        tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then   'checking for the same ticker
       
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value   'adding the volumes
            
          End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
           'comparing the ticker in the previous row and the current row to find the starting
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          
          End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        'comparing the ticker in the next row and the current row to find the ending
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                              
            '3d) Increase the tickerIndex.
             tickerIndex = tickerIndex + 1 'tickerIndex is increased as the next row's ticker doesnot match the previous row's ticker
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate    'Activating the output worksheet
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
## Results
The output Analysis of the stock data for the year 2017 is


![2017](https://user-images.githubusercontent.com/95719819/148635207-2e9975d9-c482-4d59-9ac9-5d0d7d6959a9.PNG)

The output Analysis of the stock data for the year 2018 is


![2018](https://user-images.githubusercontent.com/95719819/148635264-a0e506b4-5113-42d9-98cb-57768938e360.PNG)

## Summary
#### All Stock Analysis Refactoring
> The Refactored code runs faster than the original VBA script. The below images shows the significant change in the run time of the original code and the refactored code.
 
 The run time of a **non-refactored code** for the year **2017** is **0.9160156 seconds**.  

![2017_Time](https://user-images.githubusercontent.com/95719819/148636123-a0a137c2-c2f4-41b5-8c2b-88fb30b970b6.PNG)

The run time of the **refactored code** for the year **2017** is **0.1757813 seconds**.  

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95719819/148636136-1c01e087-8298-4f8e-8980-0033904ab9a6.png)

The run time of a **non-refactored code** for the year **2018** is **0.942328 seconds**.  

![2018_Time](https://user-images.githubusercontent.com/95719819/148636228-64d4aae4-aa33-4f10-83e6-1427c7f167bc.PNG)

The run time of the **refactored code** for the year **2018** is **0.1972656 seconds**.  

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95719819/148636236-ef0952ce-a7ac-4d1d-94bc-80c90ca6fa1d.PNG)  

> There is no significant disadvantage in refactoring the code for analysing the given stock data.

#### General Refactoring

> The advantage of refactoring a code leads to better quality which will be easily readable and runs faster.  
> The potential disadvantage is it is highly risky when the application is so big, where you might have to retest lots of functionalities for the bugs which is time consuming.



