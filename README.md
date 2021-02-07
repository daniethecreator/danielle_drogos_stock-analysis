# danielle_drogos_stock-analysis
## Module 2 Challenge Stock Analysis VBA

### Overview
The purpose of this project was to at first look at the stock "DQ" and then see if it yielded a high return for Steve's parents.
After looking at "DQ", there were 12 other stocks that Steve wanted to look into with a formula that automatically calculated the volume and return for each year. 
Since the initial code set up did run successfully, the new task was to see if there was a way to rewrite or refactor the code in order to improve efficiency. 

### Results
The refactored code retains all the functionality of the original code, including the majority of the initial set-up. First, startTime and endTime were set to the "Single" data type (Dim startTime As Single)(Dim endTime as Single).  Then the year input box was added so that the user could provide which year they would like to run the analysis on. (yearValue = InputBox("What year would you like to run the analysis on?")).
Second, the timer was started in order to track the run time of the code (starTime = Timer).
The next step was to format the output on the 'All Stocks Analysis tab' by first activating the tab that needed to be used (Worksheets("All Stocks Analysis").Activate) and then adding the header for whichever year input was added. (Range("A1").Value = "All Stocks (" + yearValue + ")".
Additional column headers were added next, Cells(3,1).Value = "Ticker" Cells(3,2). Value = "Total Daily Volume. Cells(3,3). Value = "Return"
After the analysis sheet was set up visually, the 12 tickers needed to be hard coded. This started with Dim tickers(12) As String, and then numbers 0-11 were assigned one of the 12 stocks Steve wanted to look at. (tickers(0)= "AY") tickers(1) = "CSIQ") etc...

After the variable "tickers" was defined, the worksheet for the year provided from user input was activated (Worksheets(yearValue).Activate)
Code was then added to get the number of rows to loop over (RowCount = Cells(rows.Count,"A"). End(xlUp).row)

The next line is where the refactored code starts to differ from the original code. In the refactored code, the variable tickerIndex was created (Dim tickerIndex As Integer) and set to 0 (tickerIndex = 0), versus the original code which started looping through the tickers (For i = 0 To 11) (ticker = tickers(i)).

Along with the new tickerIndex variable, three arrays were additionally set up (Dim tickerVolumes(12) As Long), (Dim tickerStartingPrices(12) As Single), and (Dim tickerEndingPrices (12) As Single).
The tickerVolumes array was to store the total volume of each of the tickers
The tickerStartingPrices array was to store the starting price for each of the tickers
The tickerEndingPrices array was to store the ending price for each of the tickers

With the index and arrays set up for the refactored code, the tickerVolumes array was initialized to 0
For j =0 to 11
    ticker Volumes(j) = 0

Then to loop over all the rows, a for loop was set to start at 2 and increment by one to the RowCount variable value set up earlier
For i = 2 to RowCount

Now, the code was ready to get the information for each ticker and each array. First, it looked to get the total volume by grabbing each cell of data and adding it to itself for the current ticker using an If Then statement to determine whether to the current row had the current ticker. 
 If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       End If

Continuing to use an If Then statement, it looked at the Starting Price by determining when the pevious cell had a different value than the current cell in the ticker column. If it did, the value was stored for starting price.
 If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If

Similarly, for the Ending Pice, the code checked to see if the subsequent cell was different, and if it was, it was stored. 
  If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

The next step was to make sure it got the volume, starting price, and ending price for each ticker, to do this it was told to increase by 1 up to 11 which was set earlier. 
tickerIndex = tickerIndex + 1

Lastly, the code needed to print the values it obtained from each ticker from the earlier steps in the Analysis worksheet.
 For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

 After all the data was collected and printed for the 12 stocks, there was some formatting additions and to color the return red or green depending if it was positive or negative.       
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
Finally, the timer was ended with a message box that showed how long it took to run the code. 
endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

### Summary
