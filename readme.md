# VBA of Wall Street Challenge

## Purose
'The purpose and background are well defined (2 pt).

The purpose of this challenge was to refactor code we wrote during the asynchronous section to do the same calculations and formating but in a more efficent manor to save time and preserve system resources. The code we wrote was to analyse stocks for our friend Steve, a new finacial planner putting together an investment strategy for his parents. The data set he asked up to look at contains information on prices and trading volumes of twelve stocks in the green energy sector for the years 2017 and 2018. Steve says a summary of the starting prices and ending prices along with the trading volume for each year will help him make a recomendation to his parents.

## Results
### Stock Analysis
Our analysis of the data Steve provided is summarized in the following images. Each of these tables describes the starting price, ending price and total trading volume of the stocks along with the yearly return by percentage. 
![2017 Analysis](/Resources/AllStocks(2017).png)

![2018 Analysis](/Resources/AllStocks(2018).png)

The trend does not look good for this entire sector. In 2017 11 out of the 12 stocks returned profits with 4 of the 12 returning better than 100% for the year! In 2017 the entire sector was hot and it was almost impossible to lose. The numbers for 2018 tell a different story. In 2018 only 2 of the 12 stocks we tracked turned a profit. Loses were the norm for the sector with the worst proformers down as much as 60% for the year suggesting that coming out of 2017 the market thought that many of these companies were overvalued.
After looking at the trend within the sector Steve would do well to tell his parents to hold off buying into green energy until furthur analysis can be done showing that the prices withing the sector had bottomed out and were primed for a rebound.

### Code & Execution Time
Below you can see the original code I used to perform the analysis along with the refactored code. There are only a few subtile changes within the code. The original uses Nested For Loops to compile the values and output them into a new sheet before moving on to the next ticker symbol. The refactored code makes use of three dirrefent arrays to store values (tickerVolumes, tickerStartingPrices and tickerEndingPrices) eliminating the need to nest loops. This new approach cuts the execution time of the script significantly as you can see in the screenshots below the code blocks.

#### Original Code Using Nested Loops

~~~
Sub AllStocksAnalysis()

Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer


   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
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
~~~




#### Refactored Code using Arrays
~~~
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    
    'Create header row
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
    
    
    
    
    
  For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
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
~~~


#### Original Code 2017
![2017 Original](/Resources/Analysis2017.png.png

#### Refactored Code 2017
![2017 Refactor](/Resources/Refactor2017.png)

#### Original Code 2018
![2018 Original](/Resources/Analysis2018.png)

#### Refactored Code 2018
![2018 Refactor](/Resources/Refactor2018.png)

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Results
'The analysis is well described with screenshots and code (4 pt).

## Summary

The main advantage to refactoring code is to optimize it. When code is optimized it can preform the same functions and produce the same output while using fewer system resources and thus producing results faster. Refactoring code can also make more elegant and it easier to understand.

The disadvantages that must be considered when deciding weither of not to refactor some existing, working code are that it adds expense to the project, at least up front. Time and resources must be spent doing the gritty work of refactoring. It is also prossible that while attempting to refactor for ease of use a developer could break the old code, particularly if they don't understand the original well. Breaking working software in an attempt to optimize it leads lead to delays and cost overruns. 


'How do these pros and cons apply to refactoring the original VBA script?

The advantages of the refactor are fairly obvious as described in the results section. The new code ran and produced results about six times faster than the original script. The new code was able to run faster because it utilized arrays to store values for all the ticker symbols. The original code used a nested loop that printed the output values of each stock line by line before iterating to the next value. The refactored code utilizes an array to store the calulated values for each stock before outputing them all at the end of the process. It seems like a very minor change but at the run times of each script shows the use of the array made the refactored script much more efficent. 

I am having trouble coming up with any disadvantages of refactoring in this case. The only con that comes to mind is that it required a fair amount of time as I am still working to grasp the concepts and syntax of VBA. 

'There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
'There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).