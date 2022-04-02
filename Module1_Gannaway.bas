Attribute VB_Name = "Module1"
Sub MakeResult()

Dim sh As Worksheet
Dim rw As Range
Dim RowCount As Long
Dim CurrentOutRow As Integer
Dim CurrentTickerSym As String
Dim A_YrStartPrice As Double
Dim A_YrEndPrice As Double
Dim A_StockVol As Double


RowCount = 0
CurrentOutRow = 2
CurrentTickerSym = "INITIALIZE"

Set sh = ActiveSheet
For Each rw In sh.Rows

    If sh.Cells(rw.Row, 1).Value = "" Then
        Exit For
    End If
    
    If sh.Cells(rw.Row, 1).Value = "<ticker>" Then
    RowCount = RowCount - 1
    ' do label row for printouts
        sh.Cells(rw.Row, 9).Value = "Ticker"
        sh.Cells(rw.Row, 10).Value = "Yearly Change"
        sh.Cells(rw.Row, 11).Value = "Percent Change"
        sh.Cells(rw.Row, 12).Value = "Total Stock Volume"
    Else
    ' do data output row if end of <ticker> section
        If CurrentTickerSym <> sh.Cells(rw.Row, 1).Value Then
            If CurrentTickerSym <> "INITIALIZE" Then
                'print output at CurrentOutRow
                sh.Cells(CurrentOutRow, 9).Value = CurrentTickerSym
                sh.Cells(CurrentOutRow, 10).Value = A_YrEndPrice - A_YrStartPrice
                sh.Cells(CurrentOutRow, 11).Value = (A_YrEndPrice - A_YrStartPrice) / A_YrStartPrice
                sh.Cells(CurrentOutRow, 12).Value = A_StockVol
                CurrentOutRow = CurrentOutRow + 1
            End If
 
            'inititialize A_accumulators with first row of new CurrentTickerSym
            CurrentTickerSym = sh.Cells(rw.Row, 1).Value
            A_YrStartPrice = sh.Cells(rw.Row, 3).Value
            A_StockVol = sh.Cells(rw.Row, 7).Value
        Else
            'A_accumulate for a row that is NOT a new CurrentTickerSym
            A_YrEndPrice = sh.Cells(rw.Row, 6).Value  'Prepare for the situation that this is end of Year row
            A_StockVol = A_StockVol + sh.Cells(rw.Row, 7).Value
        End If

    End If

  RowCount = RowCount + 1

Next rw

MsgBox (RowCount)

End Sub

