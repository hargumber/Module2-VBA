
Sub stocks()
    'Declare variables
    
    Dim i As Long
    
    Dim ws As Worksheet
    
    'Loop through all sheets
    
    For Each ws In ThisWorkbook.Worksheets
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim QuaterlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As LongLong
        Dim Summary_table_row As Long
        Dim LastRow As Long
        
    
        'Setting starting values
            
        Summary_table_row = 2
        TotalStockVolume = 0
        'setting the first open price
        OpenPrice = ws.Cells(2, 3).Value
        'ClosePrice = 0
        'QuaterlyChange = 0
        'PercentChange = 0
            
        'Determine the last row
       
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        'Adding headers to Summary Table
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Quaterly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        
        'setting the loop
            
        For i = 2 To LastRow
            
        'checking the condition
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'Print the values?
            ws.Range("K" & Summary_table_row).Value = Ticker
            ws.Range("N" & Summary_table_row).Value = TotalStockVolume
            
            ClosePrice = ws.Cells(i, 6).Value
            
            'Calculating Quaterly change
            QuaterlyChange = ws.Cells(i, 6).Value - OpenPrice
            
            If QuaterlyChange > 0 Then
                 ws.Cells(Summary_table_row, 12).Interior.ColorIndex = 4
            End If
            
            If QuaterlyChange < 0 Then
                 ws.Cells(Summary_table_row, 12).Interior.ColorIndex = 3
            End If
            
            'Print quaterly change
            ws.Range("L" & Summary_table_row).Value = QuaterlyChange
            
            'Calculating Percent change
            PercentChange = (QuaterlyChange / OpenPrice)
                  
            'Print percent Change
            ws.Range("M" & Summary_table_row).Value = PercentChange
            
            'Format the Percentage column
            ws.Range("M" & Summary_table_row).NumberFormat = "0.00%"
            
            ws.Range("N" & Summary_table_row).NumberFormat = "0"
            
            
            Summary_table_row = Summary_table_row + 1
            TotalStockVolume = 0
            
            'Reset Open Price to the start price of next ticker
            OpenPrice = ws.Cells(i + 1, 3).Value
            
            Else
            
                Ticker = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            End If
        Next i
        
        'Header Labels
        ws.Cells(1, "Q").Value = "Ticker"
        ws.Cells(1, "R").Value = "Value"
        
        'Label for greatest % increase
        ws.Cells(2, "P").Value = "Greatest % Increase"
        
        'Range to look for percentage change
        Dim percentRange As Range
        Set percentRange = ws.Range("M:M")
                
        'Finding the greatest % increase
        ws.Cells(2, "R").Value = WorksheetFunction.Max(percentRange)
        ws.Cells(2, "R").NumberFormat = "0.00%"
        greatestIncrease = ws.Cells(2, "R").Value
        
        'Finding row corresponding to greatest increase to get Ticker
        'MsgBox (Format(greatestIncrease, "0.00%"))
        Dim greatestIncreaseRow As Range
        Set greatestIncreaseRow = percentRange.Find(What:=Format(greatestIncrease, "0.00%"), LookIn:=xlValues)
        Dim rowNum As Long
        
        rowNum = greatestIncreaseRow.Row
        
        ws.Cells(2, "Q").Value = ws.Cells(rowNum, "K").Value
        
        'Label for greatest % decrease
        ws.Cells(3, "P").Value = "Greatest % Decrease"
        
        'Finding the greatest % decrease
        greatestDecrease = WorksheetFunction.Min(percentRange)
        ws.Cells(3, "R").Value = greatestDecrease
        ws.Cells(3, "R").NumberFormat = "0.00%"
        
        'Finding row corresponding to greatest decrease to get Ticker
        Dim greatestDecreaseRow As Range
        Set greatestDecreaseRow = percentRange.Find(What:=Format(greatestDecrease, "0.00%"), LookIn:=xlValues)
        
        rowNum = greatestDecreaseRow.Row
        
        ws.Cells(3, "Q").Value = ws.Cells(rowNum, "K").Value
        
        'Label for greatest total volume
        ws.Cells(4, "P").Value = "Greatest Total Volume"
        
        'Finding the greatest total volume
        Dim volumeRange As Range
        Set volumeRange = ws.Range("N:N")
        greatestVolume = WorksheetFunction.Max(volumeRange)
        ws.Cells(4, "R").Value = greatestVolume
        
        'Finding row corresponding to greatest volume to get Ticker
        Dim greatestVolumeRow As Range
        Set greatestVolumeRow = volumeRange.Find(What:=greatestVolume, LookIn:=xlValues)
        
        rowNum = greatestVolumeRow.Row
        
        ws.Cells(4, "Q").Value = ws.Cells(rowNum, "K").Value
        
    Next ws
    
End Sub
