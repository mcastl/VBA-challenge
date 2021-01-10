Attribute VB_Name = "Module1"
Sub Stocks()

' Insert Text for each Worksheet
For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greates % Decrease"
    ws.Range("O4").Value = "Greatest Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

' Declare Variables
Dim x, iRow, PriceTrack As Long
Dim ticker As String
Dim OpenVal, CloseVal, YChg, PChg, StVol As Double
Dim TickerTrack As Integer
    
' Declare variables for looping
TickerTrack = 2
PriceTrack = 2
StVol = 0

'Set up the count as the number of filled rows in column "A"
iRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop for tickers and match with their volumes
For x = 2 To iRow
   If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
      ticker = ws.Cells(x, 1).Value
      StVol = StVol + ws.Range("G" & x).Value
      ' Print values
      ws.Range("I" & TickerTrack).Value = ticker
      ws.Range("L" & TickerTrack).Value = StVol
                
      'Compute Yearly Change
      OpenVal = ws.Range("C" & PriceTrack).Value
      CloseVal = ws.Range("F" & x).Value
      YChg = CloseVal - OpenVal
      ws.Range("J" & TickerTrack).Value = YChg
      ' Conditional formatting with positive chg in green and negative in red
        If ws.Range("J" & TickerTrack).Value < 0 Then
           ws.Range("J" & TickerTrack).Interior.ColorIndex = 3
        Else
           ws.Range("J" & TickerTrack).Interior.ColorIndex = 4
        End If
                
        'Compute Percent Change
        If OpenVal = 0 Then
            PChg = 0
        Else
            PChg = YChg / OpenVal
            ws.Range("K" & TickerTrack).Value = PChg
        End If
          
      'Reset
      TickerTrack = TickerTrack + 1
      PriceTrack = x + 1
      StVol = 0
    Else
      StVol = StVol + ws.Range("G" & x).Value
    End If
Next x
    
' BONUS
' Declare variables
Dim GtstIn, GtstDc, GtstSt As Double
GtstIn = 0
GtstDc = 0
GtstSt = 0

'Set up the count as the number of filled rows in column "I"
pRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To pRow
    If ws.Range("K" & i).Value > GtstIn Then
        GtstIn = ws.Range("K" & i).Value
        ws.Range("Q2").Value = GtstIn
        ws.Range("P2").Value = ws.Range("I" & i)
    ElseIf ws.Range("K" & i).Value < GtstDc Then
        GtstDc = ws.Range("K" & i).Value
        ws.Range("Q3").Value = GtstDc
        ws.Range("P3").Value = ws.Range("I" & i)
    ElseIf Range("L" & i).Value > GtstSt Then
        GtstSt = ws.Range("L" & i).Value
        ws.Range("Q4").Value = GtstSt
        ws.Range("P4").Value = ws.Range("I" & i)
        
    End If
Next i

' Format cells
ws.Range("K:K, Q2:Q3").NumberFormat = "0.00%"
ws.Rows("1").HorizontalAlignment = xlCenter
ws.Rows("1").Font.Bold = True
ws.Range("O2:O4").Font.Bold = True
ws.Columns("A:Q").EntireColumn.AutoFit
ws.Columns("H").ColumnWidth = 4
ws.Columns("M:N").ColumnWidth = 2

Next ws

End Sub
