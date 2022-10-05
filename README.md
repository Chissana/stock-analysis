# stock-analysis

## Overview of Project

### Purpose
Analyzing and refactoring stock data using VBA so we loop through the data one time and collect all the information. We will also display the time it takes our code to process the given information. 

## Results

### Analysis with Screenshots
The stock performance between 2017 and 2018 mostly differ by their percent of return. They are almost polar opposite with 2018 having a mostly negative percent of return and 2017 having a mostly postive percent of return. The Total Daily Volume of each stock is also drastically different in 2017 then it was in 2018. 
<img width="1680" alt="Screen Shot 2022-10-04 at 10 37 51 PM" src="https://user-images.githubusercontent.com/113744353/193970317-8e49cce3-9c90-4579-966a-37263afb5b5a.png">

<img width="1680" alt="Screen Shot 2022-10-04 at 10 39 46 PM" src="https://user-images.githubusercontent.com/113744353/193970332-39ffde7d-97fe-4321-acff-7acf8a64a1c5.png">

The execution times of the refactored script seems to be the same for both the years 2017 and 2018.
<img width="1680" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/113744353/193970669-f516333d-98d7-471d-bfcc-a26b48ed62dc.png">

<img width="1680" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/113744353/193970683-712da4f0-763d-4303-9e02-984d8d3aac85.png">

The execution times of the original script show that 2017 ran faster than 2018.
<img width="1680" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/113744353/193971584-6ad2b98b-3bba-4038-b6f7-e3c761cf01f4.png">
<img width="1680" alt="VBA_Challenge_2018 10 41 36 PM" src="https://user-images.githubusercontent.com/113744353/193971614-2b1fefc1-8ab4-40f6-9aa4-8262d3c64f63.png">

### Analysis with Code 
After I refactored the code certain variables were added. Specifically tickerIndex which was a big part of the code because every time it looped we increased the tickerIndex by one as shown below:

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
       '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
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
