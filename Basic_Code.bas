Attribute VB_Name = "Basic_Code"
Sub BasicSummary()

'---------------------Populate basic summary table---------------------

Application.ScreenUpdating = False


For Each ws In Worksheets

    Dim Ticker As String
    Dim YOP, YCP, YearlyChange, PctChg As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Ticker_Vol_Total As Double
    Ticker_Vol_Total = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YCP"
    ws.Cells(1, 11).Value = "YOP"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Precent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    
        
    For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
                
            YCP = ws.Cells(i, 6).Value
            
            Ticker_Vol_Total = Ticker_Vol_Total + ws.Cells(i, 7).Value
                
            ws.Range("I" & Summary_Table_Row).Value = Ticker
                      
            ws.Range("J" & Summary_Table_Row).Value = YCP
            
            ws.Range("N" & Summary_Table_Row).Value = Ticker_Vol_Total
                
            Summary_Table_Row = Summary_Table_Row + 1
            
            Ticker_Vol_Total = 0
            
        Else

            Ticker_Vol_Total = Ticker_Vol_Total + ws.Cells(i, 7).Value
                      
        End If
         
    Next i
    
    
    Summary_Table_Row = 2


    For i = 1 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            YOP = ws.Cells(i + 1, 3).Value

            ws.Range("K" & Summary_Table_Row).Value = YOP

            Summary_Table_Row = Summary_Table_Row + 1

        End If

    Next i


    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row


    For i = 2 To LastRow

        ws.Cells(i, 12).Value = Format(ws.Cells(i, 10).Value - ws.Cells(i, 11).Value, "Standard")

        If ws.Cells(i, 11).Value = 0 Then
            ws.Cells(i, 13).Value = 0
        Else
            ws.Cells(i, 13).Value = Format(ws.Cells(i, 12) / ws.Cells(i, 11).Value, "Percent")
         End If

    Next i
    
    
    For r = 2 To LastRow
        If ws.Cells(r, 12).Value < 0 Then
            ws.Cells(r, 12).Interior.ColorIndex = 3
        Else
            ws.Cells(r, 12).Interior.ColorIndex = 4
        End If
    Next r
    
    ws.Columns("J:K").Delete
    
    ws.Columns("I:L").AutoFit
    
Next ws
    
Application.ScreenUpdating = True

MsgBox ("Basic Summaries Created")
    
End Sub
    
    
    




