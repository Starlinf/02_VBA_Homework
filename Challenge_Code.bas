Attribute VB_Name = "Challenge_Code"

Sub ChallengeSummary()

Application.ScreenUpdating = False

For Each ws In Worksheets

    Dim MaxTicker, MinTicker As String
    Dim Max, Min, GVol As Double

    Dim Challenge_Summary_Table_Row As Integer
    Challenge_Summary_Table_Row = 2


    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"


    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"


    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row


    'Find Greatest % Increase

    Max = 0

    For i = 2 To LastRow

        If ws.Cells(i, 11).Value > Max Then

            Max = ws.Cells(i, 11).Value

            MaxTicker = ws.Cells(i, 9).Value

            ws.Range("P" & Challenge_Summary_Table_Row).Value = Format(Max, "Percent")

            ws.Range("O" & Challenge_Summary_Table_Row).Value = MaxTicker

        End If

    Next i


    'Find Min % Decrease

    Challenge_Summary_Table_Row = 3


    Min = 0

    For i = 2 To LastRow

        If ws.Cells(i, 11).Value < Min Then

            Min = ws.Cells(i, 11).Value

            MinTicker = ws.Cells(i, 9).Value

            ws.Range("P" & Challenge_Summary_Table_Row).Value = Format(Min, "Percent")

            ws.Range("O" & Challenge_Summary_Table_Row).Value = MinTicker

        End If

    Next i


    'Find Greatest Total Volumne

    Challenge_Summary_Table_Row = 4

    Max = 0

    For i = 2 To LastRow

        If ws.Cells(i, 12).Value > Max Then

            Max = ws.Cells(i, 12).Value

            MaxTicker = ws.Cells(i, 9).Value

            ws.Range("P" & Challenge_Summary_Table_Row).Value = Max

            ws.Range("O" & Challenge_Summary_Table_Row).Value = MaxTicker

        End If

    Next i

    ws.Range("N:P").Columns.AutoFit
    
Next ws


Application.ScreenUpdating = True

MsgBox ("Challenge Summaries Created")


End Sub


