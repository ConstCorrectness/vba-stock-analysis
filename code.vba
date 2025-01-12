Sub RunMacroOnAllSheets()
    Dim ws As Worksheet
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        Calc
    Next ws
End Sub

Sub Calc()
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim percentageChange As Double
    Dim i As Long

    ' Variables for greatest values
    Dim greatestIncrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolume As Double
    Dim greatestVolumeTicker As String

    ' Initialize greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' Set the active worksheet
    Set ws = ActiveSheet
    ' Find the last row in the data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Add headers for the summary table
    ws.Cells(1, 9).Value = "Ticker"             ' Column I
    ws.Cells(1, 10).Value = "Quarterly Change"  ' Column J
    ws.Cells(1, 11).Value = "Percentage Change" ' Column K
    ws.Cells(1, 12).Value = "Total Volume"      ' Column L
    summaryRow = 2 ' Start row for the summary table

    startRow = 2 ' Start of the first ticker block

    ' Loop through each row of data
    For i = 2 To lastRow
        ' Check if the ticker changes (or it's the last row)
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
            ' Capture the ticker
            ticker = ws.Cells(startRow, 1).Value

            ' Determine start and end rows for the ticker
            endRow = i

            ' Calculate opening price, closing price, and total volume
            openingPrice = ws.Cells(startRow, 3).Value
            closingPrice = ws.Cells(endRow, 4).Value
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 5), ws.Cells(endRow, 5)))

            ' Calculate percentage change
            percentageChange = 0
            If openingPrice <> 0 Then
                percentageChange = ((closingPrice - openingPrice) / openingPrice) * 100
            End If

            ' Output results to the summary table
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = closingPrice - openingPrice
            ws.Cells(summaryRow, 11).Value = percentageChange
            ws.Cells(summaryRow, 12).Value = totalVolume

            ' Apply conditional formatting to Column J (Quarterly Change)
            Dim cell As Range
            Set cell = ws.Cells(summaryRow, 10) ' Column J
            If cell.Value > 0 Then
                cell.Interior.Color = RGB(144, 238, 144) ' Light green
            ElseIf cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 99, 71) ' Light red
            Else
                cell.Interior.ColorIndex = xlNone ' No fill
            End If

            ' Track greatest percentage increase
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = ticker
            End If

            ' Track greatest percentage decrease
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = ticker
            End If

            ' Track greatest total volume
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If

            ' Move to the next summary row
            summaryRow = summaryRow + 1

            ' Reset start row for the next ticker
            startRow = i + 1
        End If
    Next i

    ' Output the greatest values below the summary table
    ws.Cells(summaryRow + 2, 9).Value = "Greatest % Increase"
    ws.Cells(summaryRow + 2, 10).Value = greatestIncreaseTicker
    ws.Cells(summaryRow + 2, 11).Value = greatestIncrease



    ws.Cells(summaryRow + 3, 9).Value = "Greatest % Decrease"
    ws.Cells(summaryRow + 3, 10).Value = greatestDecreaseTicker
    ws.Cells(summaryRow + 3, 11).Value = greatestDecrease



    ws.Cells(summaryRow + 4, 9).Value = "Greatest Total Volume"
    ws.Cells(summaryRow + 4, 10).Value = greatestVolumeTicker
    ws.Cells(summaryRow + 4, 11).Value = greatestVolume
End Sub



