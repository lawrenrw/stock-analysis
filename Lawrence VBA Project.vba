Attribute VB_Name = "Module1"
Sub Module_2_Challenge()
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long, Summary_Table_Row As Integer
    Dim Ticker As String
    Dim Ticker_Total As Double
    Dim openingPrice As Double, closingPrice As Double
    Dim priceChange As Double, priceChangePercent As Double

    ' PART 6 variables
    Dim maxValue As Double, minValue As Double, maxVolume As Double
    Dim maxIncreaseRow As Long, maxDecreaseRow As Long, maxVolumeRow As Long
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String

    ' Loop through each worksheet
    For Each ws In Worksheets
        ' Initialize summary table headers
        With ws
            .Cells(1, 9).Value = "Ticker Type"               ' Column I
            .Cells(1, 10).Value = "Quarterly Change"         ' Column J
            .Cells(1, 11).Value = "Percent Change"           ' Column K
            .Cells(1, 12).Value = "Total Stock Volume"       ' Column L
            .Cells(2, 15).Value = "Greatest % Increase"      ' Column P
            .Cells(3, 15).Value = "Greatest % Decrease"      ' Column P
            .Cells(4, 15).Value = "Greatest Total Volume"    ' Column P
            .Cells(1, 16).Value = "Ticker"                   ' Column Q
            .Cells(1, 17).Value = "Value"                    ' Column R
        End With

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Ticker_Total = 0

        ' Loop through rows to process tickers
        For i = 2 To lastRow
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value ' Column G: Volume

            ' Check for new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value

                ' Identify opening price (first occurrence in Column C) and closing price (last occurrence in Column F)
                openingPrice = ws.Cells(Application.WorksheetFunction.Match(Ticker, ws.Range("A2:A" & lastRow), 0) + 1, 3).Value ' First occurrence of Column C
                closingPrice = ws.Cells(i, 6).Value          ' Last occurrence of Column F for the ticker

                ' Calculate quarterly change and percent change
                priceChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    priceChangePercent = priceChange / openingPrice ' No multiplication by 100 here
                Else
                    priceChangePercent = 0
                End If

                ' Write results to the summary table
                With ws
                    .Cells(Summary_Table_Row, 9).Value = Ticker           ' Column I: Ticker
                    .Cells(Summary_Table_Row, 10).Value = priceChange     ' Column J: Quarterly Change
                    .Cells(Summary_Table_Row, 11).Value = priceChangePercent ' Column K: Percent Change
                    .Cells(Summary_Table_Row, 12).Value = Ticker_Total    ' Column L: Total Volume
                End With

                ' Reset Ticker Total and increment summary table row
                Ticker_Total = 0
                Summary_Table_Row = Summary_Table_Row + 1
            End If
        Next i

        ' Apply formatting for Column K (Percent Change) as a percentage
        With ws.Range("K2:K" & Summary_Table_Row - 1)
            .NumberFormat = "0.00%" ' Format as percentage with two decimal places
        End With

        ' Apply conditional formatting for the quarterly change column (Column J)
        With ws.Range("J2:J" & Summary_Table_Row - 1)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative
        End With

        ' PART 6: Find greatest increase, decrease, and total volume
        With ws
            lastRow = .Cells(.Rows.Count, 11).End(xlUp).Row ' Percent Change (Column K)
            maxValue = Application.WorksheetFunction.Max(.Range("K2:K" & lastRow))
            minValue = Application.WorksheetFunction.Min(.Range("K2:K" & lastRow))
            maxVolume = Application.WorksheetFunction.Max(.Range("L2:L" & lastRow)) ' Total Volume (Column L)

            maxIncreaseRow = Application.WorksheetFunction.Match(maxValue, .Range("K2:K" & lastRow), 0) + 1
            maxDecreaseRow = Application.WorksheetFunction.Match(minValue, .Range("K2:K" & lastRow), 0) + 1
            maxVolumeRow = Application.WorksheetFunction.Match(maxVolume, .Range("L2:L" & lastRow), 0) + 1

            maxIncreaseTicker = .Cells(maxIncreaseRow, 9).Value ' Column I: Ticker
            maxDecreaseTicker = .Cells(maxDecreaseRow, 9).Value ' Column I: Ticker
            maxVolumeTicker = .Cells(maxVolumeRow, 9).Value     ' Column I: Ticker

            ' Write results to summary table
            .Cells(2, 16).Value = maxIncreaseTicker
            .Cells(2, 17).Value = maxValue
            .Cells(3, 16).Value = maxDecreaseTicker
            .Cells(3, 17).Value = minValue
            .Cells(4, 16).Value = maxVolumeTicker
            .Cells(4, 17).Value = maxVolume
        End With

        ' Format Q2 and Q3 as percentages
        With ws.Range("Q2:Q3")
            .NumberFormat = "0.00%" ' Format Q2 and Q3 as percentages
        End With

    Next ws

    MsgBox "Processing Complete!", vbInformation
End Sub

