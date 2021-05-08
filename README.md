# VBAchallenge

VBA CHallenge - The VBA of Wall Street


Sub VBAStock()

' Set Dimensions
Dim total As Double
Dim i As Long
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Single
Dim ws As Worksheet

For Each ws In Worksheets

    'Set values for each worksheet
    j = 0
    total = 0
    change = 0
    start = 2
    
' Set title row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"



'Get row number of last row with data

rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    'If ticker changes then print results
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Store Results in Variables
            total = total + ws.Cells(i, 7).Value
            
            'Handle zero total volume
            If total = 0 Then
            
            'Print the results
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + j).Value = 0
            ws.Range("K" & 2 + j).Value = "%" & 0
            ws.Range("L" & 2 + j).Value = 0
        Else
        
            'Find first non-zero starting value
            If ws.Cells(start, 3) = 0 Then
                For find_value = start To i
                    If ws.Cells(find_value, 3).Value <> 0 Then
                        start = find_value
                        Exit For
                        
                    End If
                Next find_value
            End If
            
            'Calculate change
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
            
            'Start of the next stock ticker
            start = i + 1
            
            'Print the results
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + j).Value = Round(change, 2)
            ws.Range("K" & 2 + j).Value = "%" & percentChange
            ws.Range("L" & 2 + j).Value = total
            
            'Colors positive green and negative red
            Select Case change
                Case Is > 0
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 4
                
                Case Is < 0
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 3
                    
                Case Else
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 0
                End Select
             End If
             
             
             'Reset variables for new atock ticker
             total = 0
             change = 0
             j = j + 1
             
        'If ticker is still the same ad results
        Else
            total = total + ws.Cells(i, 7).Value
            
        End If
        
Next i
Next ws


End Sub
