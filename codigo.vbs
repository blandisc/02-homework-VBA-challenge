Sub stock_analysis()

Dim i, wsNum, C, porcentageIncrease As Integer

Dim R, t, rowdelimitier, volume, current, openPrice, closePrice, finalPrice, rowdelimitier2, greatestDecrease, greatestIncrease, greatestVolume As Long

Dim ws As Worksheet

For Each ws In Worksheets

ws.Activate

rowdelimitier = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Establish Row Headers

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Establish loosers and winners
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'Determine openPrice and current
current = 2

openPrice = Cells(2, 3).Value

For R = 2 To rowdelimitier

For C = 1 To 1

ws.Cells(R, C).Select

'Send Message

If ws.Cells(R + 1, C).Value <> ws.Cells(R, C).Value Then

MsgBox (ws.Cells(R, C).Value)

End If

'Sum and print volume

If ws.Cells(R, 1).Value = ws.Cells(R + 1, 1).Value Then

volume = ws.Cells(R, 7).Value + volume

ElseIf ws.Cells(R + 1, C).Value <> ws.Cells(R, C).Value Then

volume = ws.Cells(R, 7).Value + volume
ws.Cells(current, 9).Value = ws.Cells(R, 1)
ws.Cells(current, 12).Value = volume
volume = 0

End If

'Calculate Yearly Change

If ws.Cells(R + 1, C).Value <> ws.Cells(R, C).Value Then
finalPrice = openPrice - ws.Cells(R, 6).Value
ws.Cells(current, 10).Value = finalPrice

'Calculate % increase or decrease
If openPrice = 0 Then
increase = finalPrice
ws.Cells(current, 11).Value = increase
ws.Cells(current, 11).Value = ws.Cells(current, 11).Value & " %"

Else
increase = (finalPrice / openPrice)
ws.Cells(current, 11).NumberFormat = "0.00%"
ws.Cells(current, 11).Value = increase
End If


'Give formatting, reset openPrice and move current

openPrice = ws.Cells(R + 1, 3).Value

End If

If increase > 0 Then
ws.Cells(current, 10).Interior.ColorIndex = 10
increase = 0
current = current + 1

End If
    
If increase < 0 Then
ws.Cells(current, 10).Interior.ColorIndex = 3
increase = 0
current = current + 1
End If

    Next C
        Next R
        
minPercent = WorksheetFunction.Min(Range(["K2:K20000"]))

For R = 2 To rowdelimitier


    If ws.Cells(R + 1, 11).Value > Cells(R, 11) Then
    maxPercent = ws.Cells(R + 1, 11).Value
    tickerMax = ws.Cells(R + 1, 9).Value
    End If
    
If ws.Cells(R, 11).Value = minPercent Then
    tickerMin = ws.Cells(R, 9).Value
End If
                
            
If ws.Cells(R + 1, 12).Value > Cells(R, 12) Then
    maxVolume = ws.Cells(R + 1, 12).Value
    tickerVol = ws.Cells(R + 1, 9).Value
End If
                
    Next R
    
                
ws.Range(["O2"]).Value = tickerMax
ws.Range(["P2"]).Value = maxPercent
ws.Range(["P2"]).NumberFormat = "0.00%"

ws.Range(["O3"]).Value = tickerMin
ws.Range(["P3"]).Value = minPercent
ws.Range(["P3"]).NumberFormat = "0.00%"

ws.Range(["O4"]).Value = tickerVol
ws.Range(["P4"]).Value = maxVolume

Next ws

End Sub
