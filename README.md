# VBAchallenge

VBA CHallenge - The VBA of Wall Street

Sub easyOption()

' Set Dimensions
Dim ws As Worksheet
Dim total As Double
Dim j As Integer

For Each ws In Worksheets
        
        'Set variables for each sheet
        total = 0
        j = 0
        
' Get the row number of the last row with data
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

' Set title row
ws.Range("i1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

For i = 2 To RowCount
        'If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Print ticker symbol
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
            
            'Print total
            ws.Range("j" & 2 + j).Value = total
            
            'Rest total
            total = 0
            
            'move to next row
            j = j + 1
             
        'Else keep adding to the total volume
        Else
            total = total + Cells(i, 7).Value
        End If
            
    Next i
    
Next ws

End Sub
