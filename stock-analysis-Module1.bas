
Sub stock()

'make executable for each worksheet

For Each ws In Worksheets



'define variables

Dim Ticker As String
Dim Summary As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim Percentage As Double
Dim Volume As Double
Dim TickerChange As Long
Dim MaxPercentage As Double
Dim MinPercentage As Double
Dim MaxVolume As Double
Dim MaxPercentageTicker As String
Dim MinPercentageTicker As String
Dim MaxVolumeTicker As String


Summary = 2
YearlyChange = 0
TickerChange = 2

'assign headers/text to specific cells

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
    
        'identify stock symbols/tickers, first open and last close annual values
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(TickerChange, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            TickerChange = (i + 1)
            Volume = Volume + ws.Cells(i, 7).Value
        
            YearlyChange = ClosePrice - OpenPrice
            Percentage = YearlyChange / OpenPrice
                If MaxPercentage < Percentage Then
                    MaxPercentage = Percentage
                    MaxPercentageTicker = Ticker
                End If
                If MinPercentage > Percentage Then
                    MinPercentage = Percentage
                    MinPercentageTicker = Ticker
                End If
                If MaxVolume < Volume Then
                    MaxVolume = Volume
                    MaxVolumeTicker = Ticker
                End If
            
            'cell formatting
            
            ws.Range("K" & Summary).NumberFormat = "0.00%"
            ws.Range("J" & Summary).NumberFormat = "0.00"
            
            'defining placements
            
            ws.Range("J" & Summary).Value = YearlyChange
            ws.Range("K" & Summary).Value = Percentage
            ws.Range("I" & Summary).Value = Ticker
            ws.Range("L" & Summary).Value = Volume
        
            Summary = Summary + 1
            Volume = 0
            
        Else
            Volume = Volume + ws.Cells(i, 7).Value
            
        'additional calcs for max and min
        
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
                
             
            
        'conditional color coding
        
        
          
            If ws.Range("J" & Summary).Value >= 0 Then
                ws.Range("J" & Summary).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Summary).Value < 0 Then
                ws.Range("J" & Summary).Interior.ColorIndex = 3
            End If
            
            If ws.Range("K" & Summary).Value > 0 Then
                ws.Range("K" & Summary).Interior.ColorIndex = 4
            ElseIf ws.Range("K" & Summary).Value < 0 Then
                ws.Range("K" & Summary).Interior.ColorIndex = 3
            End If
         End If
         
    Next i
    
            ws.Range("Q2").Value = MaxPercentage
            ws.Range("Q3").Value = MinPercentage
            ws.Range("Q4").Value = MaxVolume
            ws.Range("P2").Value = MaxPercentageTicker
            ws.Range("P3").Value = MinPercentageTicker
            ws.Range("P4").Value = MaxVolumeTicker
            
      If ws.Range("Q2").Value >= 0 Then
                ws.Range("Q2").Interior.ColorIndex = 4
            ElseIf ws.Range("Q2").Value < 0 Then
                ws.Range("Q2").Interior.ColorIndex = 3
            End If
            If ws.Range("Q3").Value >= 0 Then
                ws.Range("Q3").Interior.ColorIndex = 4
            ElseIf ws.Range("Q3").Value < 0 Then
                ws.Range("Q3").Interior.ColorIndex = 3
            End If
    Next ws
    
    
End Sub


