Sub QuarterlyAnalysis()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ws.Range("H1").Value = "Ticker"
        ws.Range("I1").Value = "Quarterly Change"
        ws.Range("J1").Value = "Percent Change"
        ws.Range("K1").Value = "Total Stock Volume"
        
        Dim openDict As Object, closeDict As Object, volumeDict As Object
        Set openDict = CreateObject("Scripting.Dictionary")
        Set closeDict = CreateObject("Scripting.Dictionary")
        Set volumeDict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = 2 To lastRow
            Dim ticker As String, dateValue As Date, openValue As Double, closeValue As Double, volume As Double
            ticker = ws.Cells(i, 1).Value
            dateValue = ws.Cells(i, 2).Value
            openValue = ws.Cells(i, 3).Value
            closeValue = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value
            
            Dim quarterKey As String
            quarterKey = ticker & "_" & Year(dateValue) & "_Q" & Application.WorksheetFunction.RoundUp(Month(dateValue) / 3, 0)
            
            If Not openDict.exists(quarterKey) Then
                openDict(quarterKey) = openValue
            End If
            closeDict(quarterKey) = closeValue
            
            If Not volumeDict.exists(ticker) Then
                volumeDict(ticker) = 0
            End If
            volumeDict(ticker) = volumeDict(ticker) + volume
        Next i
        
        Dim outputRow As Long
        outputRow = 2
        Dim key As Variant
        
        For Each key In openDict.keys
            Dim tickerKey As String
            tickerKey = Split(key, "_")(0)
            
            Dim openVal As Double, closeVal As Double, percentChange As Double, quarterlyChange As Double
            openVal = openDict(key)
            closeVal = closeDict(key)
            quarterlyChange = closeVal - openVal
            percentChange = ((closeVal - openVal) / openVal) * 100
            
            ws.Cells(outputRow, 8).Value = tickerKey
            ws.Cells(outputRow, 9).Value = quarterlyChange
            ws.Cells(outputRow, 10).Value = percentChange
            ws.Cells(outputRow, 11).Value = volumeDict(tickerKey)
            
            If quarterlyChange >= 0 Then
                ws.Cells(outputRow, 9).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(outputRow, 9).Interior.Color = RGB(255, 0, 0)
            End If
            
            If percentChange >= 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            outputRow = outputRow + 1
        Next key
        
    Next sheetName
End Sub
