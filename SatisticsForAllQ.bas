Attribute VB_Name = "SatisticsForAllQ"
Sub SatisticsForAllQuarters()

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
    QStats wsQ1
     QStats wsQ2
     QStats wsQ3
     QStats wsQ4

End Sub
Sub QStats(ws As Worksheet)
    
    ' Activate the worksheet
    ws.Activate
'"   Add functionality to your script to return the stock 'with the
'"   Greatest % increase",
'"   Greatest % decrease", and
' "  Greatest total volume"

' Construct table labels
 ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = " Ticker"
   ws.Cells(1, 17).Value = " Value"
  
 Dim QHighestPercentIncrease As Double
 Dim QLowestPercentIDecrease As Double
 
 Dim QHighestTicker As String
 Dim QLowestTicker As String
 
 Dim QHighVolume As Double
 Dim QHighVolumeTicker As String
 
 Dim LastRow As Long
 LastRow = Cells(Rows.Count, "K").End(xlUp).Row
 
  Dim i As Long
      
 QHighestPercentIncrease = ws.Cells(2, 11).Value
 QLowestPercentIDecrease = ws.Cells(2, 11).Value
 QHighVolume = ws.Cells(2, 12).Value
 
  For i = 3 To LastRow
        If ws.Cells(i, 11).Value > QHighestPercentIncrease Then
                            QHighestPercentIncrease = ws.Cells(i, 11).Value
                            QHighestTicker = ws.Cells(i, 9).Value
                     ElseIf Cells(i, 11).Value < QLowestPercentIDecrease Then
                          QLowestPercentIDecrease = ws.Cells(i, 11).Value
                          QLowestTicker = ws.Cells(i, 9).Value
        End If
                          
        If ws.Cells(i, 12).Value > QHighVolume Then
                            QHighVolume = ws.Cells(i, 12).Value
                            QHighVolumeTicker = ws.Cells(i, 9).Value
                            
        End If
  Next i
  
  
   ws.Cells(2, 16).Value = QHighestTicker
   ws.Cells(2, 17).Value = QHighestPercentIncrease
  
   ws.Cells(3, 16).Value = QLowestTicker
  ws.Cells(3, 17).Value = QLowestPercentIDecrease
  
 ws.Cells(4, 16).Value = QHighVolumeTicker
 ws.Cells(4, 17).Value = QHighVolume
 'Autofit the columns
Columns("O:Q").AutoFit
 
  End Sub

