Attribute VB_Name = "Module1"
Sub AnalyzeAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim i As Long
    Dim outputRow As Long
    Dim greatestIncreaseTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        greatestIncrease = -1E+30 ' Set to a very low number
        greatestDecrease = 1E+30 ' Set to a very high number
        greatestVolume = 0 ' Initializing for the highest volume
        
        ' Find the last row in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set the first row for output
        outputRow = 2
        
        ' Initialize the first values
        volume = 0
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if the next ticker is different, indicating the end of the current ticker's data
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                
                ' Add the volume for this ticker
                volume = volume + ws.Cells(i, 7).Value
                                
                Dim quarterlyChange As Double
                Dim percentChange As Double
                
                quarterlyChange = closePrice - openPrice
                percentChange = (quarterlyChange) / openPrice
                
                ' Output the results to the specified columns
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = volume
                
                ' Check for the greatest % increase
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                ' Check for the greatest % decrease
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                ' Check for the greatest total volume
                If volume > greatestVolume Then
                    greatestVolume = volume
                    greatestVolumeTicker = ticker
                End If
                
                ' Move to the next row for output
                outputRow = outputRow + 1
                
                ' Reset the volume for the next ticker
                volume = 0
                
                ' Set the opening price for the next ticker (if it's not the last row)
                If i + 1 <= lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                ' Add the volume for the same ticker
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output the results to the respective sheet (all sheets)
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
    Next ws
    
    MsgBox ("Analysis completed for all sheets!")
End Sub

