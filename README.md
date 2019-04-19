# VBA-_HW_Easy

Sub stock_analysis()
'delcare varables and make sure to activate the WS

Dim ws As Worksheet
Dim ticker As String
Dim stock_vol As Double
Dim summary As Integer

For Each ws In Worksheets
ws.Activate

'set stock volume as 0

stock_vol = 0

'keep track of summary table
summary = 2
    
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"
    
 ' Determine the Last Row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
        For i = 2 To lastrow

            ' check is we're in same value or not, if it is not...
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' add ticker header
                ticker = Cells(i, 1).Value

                ' add to stock vol
                stock_vol = stock_vol + Cells(i, 7).Value

                'set ticker value
                ws.Range("I" & summary).Value = ticker

                'now set stock vol value
                ws.Range("J" & summary).Value = stock_vol

                ' begin count for summary
                summary = summary + 1
                
                'reset stock vol to 0 for next count.
                stock_vol = -0
            Else

                ' add to stock vol.
                Stock = Stock + Cells(i, 7).Value

            End If
              
        Next i

    Next ws
    
End Sub
