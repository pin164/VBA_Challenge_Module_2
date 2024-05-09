Attribute VB_Name = "SummaryTableForAllQuarters"
Sub SummaryTableForAllQuarters()

    Dim wsQ1 As Worksheet
    Dim wsQ2 As Worksheet
    Dim wsQ3 As Worksheet
    Dim wsQ4 As Worksheet
    
    ' Assign each specific worksheet to its corresponding variable
    Set wsQ1 = ThisWorkbook.Sheets("Q1")
    Set wsQ2 = ThisWorkbook.Sheets("Q2")
    Set wsQ3 = ThisWorkbook.Sheets("Q3")
    Set wsQ4 = ThisWorkbook.Sheets("Q4")
    
    ' Process each worksheet one by one
    QSummaryTable wsQ1
    QSummaryTable wsQ2
    QSummaryTable wsQ3
    QSummaryTable wsQ4

End Sub
Sub QSummaryTable(ws As Worksheet)
  
 ' Activate the worksheet
    ws.Activate
    
' find the last row for the data set


Dim LastRow As Long
LastRow = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' find the first and last dates for the quarter

Dim Q1stDate As Date
Dim QLastDate As Date
QLastDate = ws.Cells(LastRow, 2).Value
Q1stDate = ws.Cells(2, 2).Value

' display values for first & last date of the quarter
ws.Cells(1, 13).Value = "Notes on Quarterly "
ws.Cells(2, 13).Value = "There are " & LastRow - 1 & " entries"

ws.Cells(3, 13).Value = "1st date of Quarter = " & Q1stDate
ws.Cells(4, 13).Value = "Last date of Quarter =" & QLastDate

' dispaly summary table header row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
' set up price variables
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim QPriceChange As Double
Dim QPercentPriceChange As Double


  ' Set an initial variable for holding the share ticker
  Dim TickerName As String

  ' Set an initial variable for holding each Ticker  Volume
  Dim TickerVolume As Double
TickerVolume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim SummaryTableRow As Integer
  SummaryTableRow = 2
  

  ' Loop through all tickers
  For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        

      ' Set the Ticker name
     TickerName = ws.Cells(i, 1).Value
     
     ' if the date is the last date of the quarter save the closing price
     
            If ws.Cells(i, 2).Value = QLastDate Then
         
                ClosingPrice = ws.Cells(i, 6).Value
                
            Else
            End If
    

      ' Add to the Ticker Volume
     TickerVolume = TickerVolume + ws.Cells(i, 7).Value
     ' calculate the quarter price change
     
      QPriceChange = ClosingPrice - OpeningPrice
      
      'Calculate the percentage change
      ' calculates how much the closing price differs from the opening price, expressed as a percentage of the opening price.
 
      ' Percentage Change=(Closing Price-Opening Price)/Opening Price×100%
      
      QPercentPriceChange = QPriceChange / OpeningPrice
      
      ' Disply the Ticker in the Summary Table
      ws.Range("I" & SummaryTableRow).Value = TickerName
      ws.Range("J" & SummaryTableRow).Value = QPriceChange
      ws.Range("K" & SummaryTableRow).Value = QPercentPriceChange
      ' Display the Ticker volume to the Summary Table
      ws.Range("L" & SummaryTableRow).Value = TickerVolume
      ' display the opening and closing prices
     ' activate these two rows to check the opening and closing prices for each ticker
    '  ws.Range("O" & SummaryTableRow).Value = OpeningPrice
     ' ws.Range("P" & SummaryTableRow).Value = ClosingPrice
      

      ' Add one to the summary table row
      SummaryTableRow = SummaryTableRow + 1
      
      ' Reset the Ticker Volume
     TickerVolume = 0
    
    ' If the cell immediately following a row is the same brand...
    Else
            If ws.Cells(i, 2).Value = Q1stDate Then
    
                OpeningPrice = ws.Cells(i, 3).Value
    
            Else
           
            End If
      ' Add to the Ticker Volume
      TickerVolume = TickerVolume + Cells(i, 7).Value

    End If

  Next i
'Autofit the columns
Columns("I:M").AutoFit
'Format %
Columns("K").NumberFormat = "0.00%"
End Sub

