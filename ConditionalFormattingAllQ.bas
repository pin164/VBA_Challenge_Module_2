Attribute VB_Name = "ConditionalFormattingAllQ"
Sub ConditionalFormattingQuarterlyChangeAllQuarters()
' set up the worksheets
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
  ApplyConditionalFormattingQuarterlyChange wsQ1
  ApplyConditionalFormattingQuarterlyChange wsQ2
   ApplyConditionalFormattingQuarterlyChange wsQ3
  ApplyConditionalFormattingQuarterlyChange wsQ4
  
  
End Sub
Sub ApplyConditionalFormattingQuarterlyChange(ws As Worksheet)
   
   ' define the variables used to establish conditional formating
    Dim LastRow As Long
    Dim j As Long
  
    ' Find the last row with data in column K = Quarterly Change
    LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
      
    ' Loop through each cell  and apply conditional formatting
    For j = 2 To LastRow
        ' Check if the cell value is greater than or equal to 0
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Green color
         ' Check if the cell value is less than 0
        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0) ' Red color
                    
        End If
    Next j
End Sub
