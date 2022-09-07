
Sub StockMacro()

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"

    Dim Ticker As String
    Dim TotalStockVolume As Double

    Dim OpeningValue As Double
    Dim ClosingValue As Double
    Dim SummaryTable As Integer

    TotalStockVolume = 0
    SummaryTable = 2

    Dim YearlyChange As Double
    Dim PercentChange As Double

    OpeningValue = Cells(2, 3).Value
    ClosingValue = 3 'NEEED TO FIX'

    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            Ticker = Cells(i, 1).Value
            Range("j" & SummaryTable).Value = Ticker

            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            Range("m" & SummaryTable).Value = TotalStockVolume


            YearlyChange = ClosingValue - OpeningValue
            Range("k" & SummaryTable).Value = YearlyChange

            PercentChange = (YearlyChange / OpeningValue) * 100
            Range("l" & SummaryTable).Value = PercentChange

            SummaryTable = SummaryTable + 1

            TotalStockVolume = 0

            OpeningValue = Cells(i + 1, 3).Value

        Else
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

        End If

    Next i

    'Setting conditional formatting'

    'If YearlyChange Is Negative Then
        'Interior.ColorIndex = 7
    'Else
        'Interior.ColorIndex = 5


End Sub
