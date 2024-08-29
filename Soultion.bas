Sub StockTicker()

    Dim ws As Worksheet
    Dim total As Double
    Dim change As Double
    Dim percentchange As Double
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim greatestvolume As Double
    Dim i As Long
    Dim start As Long
    Dim rowcount As Long
    Dim j As Integer
    Dim increaseticker As String
    Dim decreaseticker As String
    Dim totalticker As String
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
    ' Add column headers across sheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Great % Increase"
    ws.Range("O3").Value = "Great % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
        
        'Set initial values
    j = 0
    total = 0
    change = 0
    start = 2
    maxincrease = -1E+308
    maxdecrease = 1E+308
    greatestvolume = -1E+308
        
        'Iterate for each row
    
    rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
    For i = 2 To rowcount
        
           
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Sum Total
    
    total = total + ws.Cells(i, 7).Value
                
        ' Percent Change
    
    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
    percentchange = (change / ws.Cells(start, 3))
                
    ' Update start variable
    
    start = i + 1
                
        ' Min/Max of Results
    If percentchange > maxincrease Then
    maxincrease = percentchange
    increaseticker = ws.Cells(i, 1).Value

End If
                
    If percentchange < maxdecrease Then
    maxdecrease = percentchange
    decreaseticker = ws.Cells(i, 1).Value

End If
                
        ' Greatest Total Volume
    If total > greatestvolume Then
    greatestvolume = total
    totalticker = ws.Cells(i, 1).Value
End If
                
        ' Print Results
    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
    ws.Range("L" & 2 + j).Value = total
    ws.Range("J" & 2 + j).Value = change
    ws.Range("J" & 2 + j).NumberFormat = "0.00"
    ws.Range("K" & 2 + j).Value = percentchange
    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                
        'Color Change
    If ws.Range("J" & 2 + j).Value < 0 Then
    ws.Range("J" & 2 + j).Interior.ColorIndex = 3

    ElseIf ws.Range("J" & 2 + j).Value > 0 Then
    ws.Range("J" & 2 + j).Interior.ColorIndex = 4

End If
                
 ' Resetting Variables

    j = j + 1
    total = 0
    change = 0       
End If
Next i
        
    ' Print Results

    ws.Range("P2").Value = increaseticker
    ws.Range("Q2").Value = maxincrease
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P3").Value = decreaseticker
    ws.Range("Q3").Value = maxdecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = totalticker
ws.Range("Q4").Value = greatestvolume
        
Next ws

End Sub



