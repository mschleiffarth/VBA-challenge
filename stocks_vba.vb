Sub stocks()

    Dim tableRow As Integer
    Dim openingValue As Double
    Dim closingValue As Double
    Dim volume As Double
    Dim WS_Count As Integer
            
    WS_Count = ActiveWorkbook.Worksheets.Count
       
    For ws = 1 To WS_Count
        ActiveWorkbook.Worksheets(ws).Cells(1, 9).Value = "ticker"
        ActiveWorkbook.Worksheets(ws).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(ws).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(ws).Cells(1, 12).Value = "Total Stock Volume"
        volume = 0
        sheetSize = ActiveWorkbook.Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
        currentSymbol = ""
    
        For I = 2 To sheetSize
                  
        
            If (currentSymbol <> ActiveWorkbook.Worksheets(ws).Cells(I, 1).Value) Then
                
                currentSymbol = ActiveWorkbook.Worksheets(ws).Cells(I, 1).Value
                lastTableRow = ActiveWorkbook.Worksheets(ws).Cells(Rows.Count, 9).End(xlUp).Row + 1
                ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 9).Value = currentSymbol
                openingValue = ActiveWorkbook.Worksheets(ws).Cells(I, 3).Value
                volume = 0
            End If
            
            volume = volume + ActiveWorkbook.Worksheets(ws).Cells(I, 7).Value
            
            If (Cells(I + 1, 1).Value <> currentSymbol) Then
                
                closingValue = ActiveWorkbook.Worksheets(ws).Cells(I, 6).Value
                ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 10).Value = closingValue - openingValue
                
                If ((closingValue - openingValue) > 0) Then
                    ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 10).Interior.ColorIndex = 4
                    
                Else
                    ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 10).Interior.ColorIndex = 3
                    
                End If
                
                
                If (openingValue <> 0) Then
                    ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 11).Value = FormatPercent((closingValue - openingValue) / openingValue)
                    
                ElseIf (closingValue = 0) Then
                    ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 11).Value = "0"
                    
                Else
                    ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 11).Value = "-"
                    
                End If
                    
                ActiveWorkbook.Worksheets(ws).Cells(lastTableRow, 12).Value = volume
                
                
            End If
            
            
        Next I
        
        ActiveWorkbook.Worksheets(ws).Columns("A:L").AutoFit
        
    Next ws
    
    

End Sub

