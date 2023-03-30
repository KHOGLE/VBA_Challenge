Attribute VB_Name = "Module1"
Sub stocks1():

'Code provided by Gina to help computer speed
Application.ScreenUpdating = False

For Each ws In Worksheets

'Column and Cell labels
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Percent Change formatting
ws.Columns(11).NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

'found on https://stackoverflow.com/questions/21554059/how-to-get-the-row-count-in-excel-vba
Dim lastrow As Long

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    k = 2

    Total = 0
    Start = 2
    Change = 0
    PerChange = 0

'This For loop was assisted by tutor, Marko Yang (As was the matching defined variables above it)
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Total = Total + ws.Cells(i, 7).Value
        Change = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
        PerChange = Change / ws.Cells(Start, 3).Value
        ws.Cells(k, 9).Value = ws.Cells(Start, 1).Value
        ws.Cells(k, 10).Value = Change
        ws.Cells(k, 11).Value = PerChange
        ws.Cells(k, 12).Value = Total
        Start = i + 1
        Total = 0
        Change = 0
        PerChange = 0
        k = k + 1
        Else
        Total = Total + ws.Cells(i, 7).Value
        End If
    Next i

'Yearly Change Conditional Formatting
    For i = 2 To lastrow
        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        'Cells(i,10).Value < 0
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
    Next i

'Greatest Values finds
Dim GPD As Double
Dim GPI As Double

    GPValue = ws.Columns(11)

    GPD = Application.WorksheetFunction.Min(GPValue)
    ws.Cells(3, 17).Value = GPD

    GPI = Application.WorksheetFunction.Max(GPValue)
    ws.Cells(2, 17).Value = GPI
        
Dim GTV As Double

    GTValue = ws.Columns(12)

    GTV = Application.WorksheetFunction.Max(GTValue)
    ws.Cells(4, 17).Value = GTV
    
'Attaching their matching ticket
'Assisted by LA Shreha and https://www.educba.com/vba-match/
Z = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & lastrow), 0)
ws.Range("P2").Value = ws.Cells(Z + 1, 9)

Y = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & lastrow), 0)
ws.Range("P3").Value = ws.Cells(Y + 1, 9)

X = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & lastrow), 0)
ws.Range("P4").Value = ws.Cells(X + 1, 9)

Next ws

'End Part of Gina's assistance
Application.ScreenUpdating = True

End Sub

