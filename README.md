# Stock Analysis
## Overview
Refactor the VBA script to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
### Purpose
The purpose of refactoring a code is to make the VBA script run faster. 
## All Stock Analysis
The stock data for the year 2017 and 2018 are given. We have to analyze the stock data to find
  - ### Total daily volume
     Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded.
  - ### Yearly return for each stock
     The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
### VBA Script
- Macro **AllStockAnalysis** is created and **startTime** and **endTime** Variables are declared with the datatype **Single**.
- The output sheet "All Stocks Analysis" is formatted as per the requirement.
- The Input box is created to get the input **yearValue** for which the analysis has to be done.
- The Header row for the out sheet is created using **Cells()** function.
- An array for all the **tickers** is initialized.
- Variables **startingPrice** and **endingPrice** are declared with the datatype double.
- Number of rows in the datasheet is calculated using **Rows.count**, to loop through the rows in the data.
- A **for** loop is created to loop through the tickers array and **totalVolume** is intialized to zero.
- A **nested for** loop is created to loop through the rows in the data.
- The **totalVolume**, **startingPrice**, and **endingPrice** are calculated after checking for conditions with the corresponding **if** statement
- The Outputs are displayed on the corresponding cells as required in the output sheet.
- Formatting of the output data is done as per the requirement by using different attributes like** FontStyle, Borders, NumberFormat, Interior.Color**, etc.
    
Sub AllStockAnalysis() 

&nbsp; &nbsp;    Dim startTime As Single  
&nbsp; &nbsp;    Dim endTime  As Single  
&nbsp; &nbsp;    Worksheets("All Stocks Analysis").Activate  
&nbsp; &nbsp;    yearValue = InputBox("What year would you like to run the analysis on?")  
&nbsp; &nbsp;    startTime = Timer  
&nbsp; &nbsp;    Range("A1").Value = "All Stocks (" + yearValue + ")"  
&nbsp; &nbsp;    Cells(3, 1).Value = "Ticker"  
&nbsp; &nbsp;    Cells(3, 2).Value = "Total Daily Volume"  
&nbsp; &nbsp;    Cells(3, 3).Value = "Return"  
&nbsp; &nbsp;    Dim tickers(11) As String   
&nbsp; &nbsp; &nbsp;     tickers(0) = "AY"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(1) = "CSIQ"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(2) = "DQ"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(3) = "ENPH"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(4) = "FSLR"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(5) = "HASI"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(6) = "JKS"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(7) = "RUN"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(8) = "SEDG"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(9) = "SPWR"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(10) = "TERP"  
&nbsp; &nbsp; &nbsp; &nbsp;     tickers(11) = "VSLR"  
&nbsp; &nbsp; &nbsp;    Dim startingPrice As Double  
&nbsp; &nbsp; &nbsp;    Dim endingPrice As Double  
&nbsp; &nbsp; &nbsp;    Sheets(yearValue).Activate  
&nbsp; &nbsp; &nbsp;    RowCount = Cells(Rows.Count, "A").End(xlUp).Row  
&nbsp; &nbsp; &nbsp;    For i = 0 To 11  
&nbsp; &nbsp; &nbsp; &nbsp;       ticker = tickers(i)  
&nbsp; &nbsp; &nbsp; &nbsp;       totalVolume = 0   
&nbsp; &nbsp; &nbsp; &nbsp;      Sheets(yearValue).Activate  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       For j = 2 To RowCount      
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;        If Cells(j, 1).Value = ticker Then  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;            totalVolume = totalVolume + Cells(j, 8).Value  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;          End If  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;            startingPrice = Cells(j, 6).Value            
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;          End If  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;            endingPrice = Cells(j, 6).Value            
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;          End If  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       Next j     
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       Worksheets("All Stocks Analysis").Activate  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       Cells(4 + i, 1).Value = ticker  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       Cells(4 + i, 2).Value = totalVolume  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1  
&nbsp; &nbsp; &nbsp;    Next i  
&nbsp; &nbsp;   Worksheets("All Stocks Analysis").Activate  
&nbsp; &nbsp; &nbsp;    Range("A3:C3").Font.FontStyle = "Bold"  
&nbsp; &nbsp; &nbsp;    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous  
&nbsp; &nbsp; &nbsp;    Range("B4:B15").NumberFormat = "#,##0"  
&nbsp; &nbsp; &nbsp;    Range("C4:C15").NumberFormat = "0.0%"  
&nbsp; &nbsp; &nbsp;    Columns("B").AutoFit  
&nbsp; &nbsp; &nbsp;    dataRowStart = 4  
&nbsp; &nbsp; &nbsp;    dataRowEnd = 15  
&nbsp; &nbsp; &nbsp;    For i = dataRowStart To dataRowEnd       
&nbsp; &nbsp; &nbsp; &nbsp;        If Cells(i, 3) > 0 Then  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            Cells(i, 3).Interior.Color = vbGreen    
&nbsp; &nbsp; &nbsp; &nbsp;        Else  
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;             Cells(i, 3).Interior.Color = vbRed  
&nbsp; &nbsp; &nbsp; &nbsp;         End If   
&nbsp; &nbsp; &nbsp;    Next i  
&nbsp; &nbsp;    endTime = Timer  
&nbsp;    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)  
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
    
    Worksheets("All Stocks Analysis").Activate      
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
   
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
 
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
    
     Worksheets(yearValue).Activate  

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim tickerIndex As Integer
    tickerIndex = 0   'tickerIndex is set to zero before iterarting over all the rows
    
    Dim tickerVolumes(12) As Long   'data type for tickerVolumes is Long
    Dim tickerStartingPrices(12) As Single   'data type for tickerStartingPrices is Single
    Dim tickerEndingPrices(12) As Single     'data type for tickersEndingPrice is Single
    
    For i = 0 To 12
        tickerVolumes(i) = 0
    Next i

    For i = 2 To RowCount
    
  
        If Cells(i, 1).Value = tickers(tickerIndex) Then   'checking for the same ticker
       
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value   'adding the volumes
            
          End If  
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          
          End If
         
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      
             tickerIndex = tickerIndex + 1 
        End If
     
    Next i
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate  
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
   
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
> The Refactored code runs faster than the original VBA script. In this analysis the potential advantage of refactoring is the reduced running time.
 The run time of a non-refactored code for the year 2017 is 0.9160156 seconds.
 ![2017_Time](https://user-images.githubusercontent.com/95719819/148636123-a0a137c2-c2f4-41b5-8c2b-88fb30b970b6.PNG)
The run time of the refactored code for the year 2017 is 0.1757813 seconds.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/95719819/148636136-1c01e087-8298-4f8e-8980-0033904ab9a6.png)
The run time of a non-refactored code for the year 2018 is 0.942328 seconds.
![2018_Time](https://user-images.githubusercontent.com/95719819/148636228-64d4aae4-aa33-4f10-83e6-1427c7f167bc.PNG)
The run time of the refactored code for the year 2018 is 0.1972656 seconds.
![VBA_Challenge_2018](https://user-images.githubusercontent.com/95719819/148636236-ef0952ce-a7ac-4d1d-94bc-80c90ca6fa1d.PNG)
#### General Refactoring
g
> Refactoring a code leads to better quality by easily readable and runs faster.  
> The potential disadvantage is it is highly risky when the application is so big, where you might have to retest lots of functionalities for the bugs which is time consuming.



