Attribute VB_Name = "Module1"
Sub stockticker()

'Establishing variables

    Dim i As Long
    Dim lastrow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim qtrChange As Double
    Dim pctChange As Double
    Dim stockVolume As Double
    Dim outputRow As Double
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncTicker As String
    Dim greatestDecTicker As String
    Dim greatestVolTicker As String
    
    Dim ws As Worksheet
    
'Worksheet Loop

    For Each ws In Worksheets

'Defining Last Row

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Initializing Variables

    openPrice = 0
    closePrice = 0
    stockVolume = 0
    pctChange = 0
    outputRow = 2
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

'Setting Headers
                
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Applying Percentage Format

    ws.Columns(11).NumberFormat = "0.00%"

'Column Loop

    For i = 2 To lastrow
    
    'Checking for first row
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
        'Applying open price value
            
            openPrice = ws.Cells(i, 3).Value
        
        End If
        
    'Checking for last row
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Writing ticker symbol
        
            ticker = ws.Cells(i, 1).Value
            ws.Cells(outputRow, 9).Value = ticker
            
        'Calculating Quarterly Change
        
            closePrice = ws.Cells(i, 6).Value
            qtrChange = closePrice - openPrice
            
        'Publishing Quarterly Change
        
            ws.Cells(outputRow, 10).Value = qtrChange
            
        'Quarterly Change Color Conditional, if greater than 0, then make green
        
            If qtrChange > 0 Then
            
                ws.Cells(outputRow, 10).Interior.ColorIndex = 4
        
        'if less than 0, then make red
                
            ElseIf qtrChange < 0 Then
                
                ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                
        'Everything else is white
        
            Else
            
                ws.Cells(outputRow, 10).Interior.ColorIndex = 2
            
            End If
            
        'Calculating Percent Change
        
            pctChange = qtrChange / openPrice
            
        'Publishing Percent Change
        
            ws.Cells(outputRow, 11).Value = pctChange
            
        'Adding Total Stock Value
        
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            
        'Publishing Total Stock Volume
        
            ws.Cells(outputRow, 12).Value = stockVolume
            
            
        'Checking for Greatest Increase
            
            If pctChange > greatestIncrease Then
            
                greatestIncTicker = ticker
                greatestIncrease = pctChange
                
            End If
            
        'Checking for Greatest Decrease
        
            If pctChange < greatestDecrease Then
            
                greatestDecTicker = ticker
                greatestDecrease = pctChange
                
            End If
            
        'Checking for Greatest Total Volume
        
            If stockVolume > greatestVolume Then
            
                greatestVolTicker = ticker
                greatestVolume = stockVolume
                
            End If
            
        'Reset Total Stock Volume
            
            stockVolume = 0
            
        'Go to the next row
            outputRow = outputRow + 1
            
        Else
        
            
        'Adding Total Stock Value
        
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    'Publish Greatest Increase, Decrease, and Volume
        
        'Creating Labels
        
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            
            
        'Publish Greatest % Increase
        
            ws.Range("O2").Value = greatestIncTicker
            ws.Range("P2").Value = greatestIncrease
            ws.Range("P2").NumberFormat = "0.00%"
            
        'Publish Greatest % Decrease
        
            ws.Range("O3").Value = greatestDecTicker
            ws.Range("P3").Value = greatestDecrease
            ws.Range("P3").NumberFormat = "0.00%"
            
        'Publish Greatest Total Volume
        
            ws.Range("O4").Value = greatestVolTicker
            ws.Range("P4").Value = greatestVolume
    
    Next ws
    
End Sub


