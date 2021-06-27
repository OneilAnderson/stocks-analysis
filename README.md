# VBA Stock Analysis
## Purpose of Analysis  

  The purpose of this analysis is to refactor our previous code so that our code can work with more efficiency. Through this, we can assist Steve into using a code that can be used for larger datasets, has a shorter time of execution and easier to understand. For this analysis,the assignment was to collect data on 12 different stocks. Once we received the ticker, total daily volume and return percentage for the year, it was placed in a chart on Excel. 
  
 ## Results
  
  To refector the code, I opened the .vbs file that was given for a layout of the code. I used this as a step-by-step guideline to successfully refactor the code. Below is the refactored code that includes a small explanation for each part of the code as well as the results from 2017 and 2018.
  
  
    '1a) Create a ticker Index

    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
      tickerVolumes(12) = tickerVolumes(i)
      tickerStartingPrices(12) = tickerStartingPrices(i)
      tickerEndingPrices(12) = tickerEndingPrices(i)
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To rowEnd
    
        '3a) Increase volume for current ticker

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

           
           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
    End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

    
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
    End If
    
            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
             
             
             tickerIndex = tickerIndex + 1
            
    End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
        
    Next i
     'Formatting data in All Stocks Analysis Sheet
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
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
              'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'displays the time in Message box
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
![All_Stocks_2017](https://user-images.githubusercontent.com/85713532/123530442-4129e500-d6c8-11eb-9bdc-64d7ccd0852a.png)
![All_Stocks_2018](https://user-images.githubusercontent.com/85713532/123530444-438c3f00-d6c8-11eb-8427-a4c632d7aed9.png)
## Summary
### Advantages/Disadvantages of Refactoring Code
  
   The advantages of refactoring code is that it makes the code cleaner and easier to understand. It also can create a faster program. For companies, I presume that refactored code can help when it comes to time consumption. Some of the disadvantages of refactoring code is that it is time comsuming and making mistakes can also occur, especially if it a large application.
   
### Advantages/ Disadvantage of Refactoring Our Original Code

   The advantages of refactoring our original code was that our code ran a faster program and is easier to understand. For both 2017 and 2018, our refactored code ran for around .3 seconds, while the original program would run longer.Below, we can see the run times for the refactored code in 2017 and 2018. The only disadvantage of refectoring this code was that it took time to do, however, I believe the pros outweighs the cons because of the increase of efficiency for the code.
  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/85713532/123531193-70445480-d6d0-11eb-8a86-78acd23d6789.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85713532/123531194-72a6ae80-d6d0-11eb-97b9-a14c672b1b7a.png)
