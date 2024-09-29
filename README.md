# Module2_VBA

'Staring module
Sub stock_data_pt1()

'Looping across worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'Creating headers for each data point collected
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Establishing Variables
Dim lastrow As Long
Dim lastrowTicker As Long
Dim i As Long
Dim Ticker_Name As String
Dim Ticker_Total As Variant
Ticker_Total = 0
Dim Summary_Table As Integer
Summary_Table = 2
Dim Openp As Double
Dim Closep As Double
Dim first As Boolean
Dim Ticker_Targets As Range
Dim Ticker_Target As Range
Dim targetcol As String
Dim outputcol As String
Dim outputcol2 As String
Dim targetstartrow As Long
Dim difference As Double
Dim maxpercent As Variant
Dim maxticker As String
Dim minpercent As Variant
Dim minticker As String
Dim maxvol As Variant
Dim maxvolticker As String

'Establishing last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Autofit columns
ws.Columns("I:Q").AutoFit

'First for loop to gather unique tickers and total volumes per unique tickers
For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker_Name = ws.Cells(i, 1).Value
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    ws.Range("I" & Summary_Table).Value = Ticker_Name
    
    ws.Range("L" & Summary_Table).Value = Ticker_Total

    Summary_Table = Summary_Table + 1

    Ticker_Total = 0

    Else

    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i

'Second for loop to compare ending and closing differences for each variable
targetstartrow = 2
lastrowTicker = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
Set Ticker_Targets = ws.Range(ws.Cells(targetstartrow, "I"), ws.Cells(lastrowTicker, "I"))
outputcol = "J"
outputcol2 = "K"

For Each Ticker_Target In Ticker_Targets
    first = False
    Openp = 0
    Closep = 0
    
        For i = 1 To lastrow
            If ws.Cells(i, "A").Value = Ticker_Target.Value Then
                If Not first Then
                    Openp = ws.Cells(i, "C").Value
                    first = True
                End If
                Closep = ws.Cells(i, "F").Value
            End If
        Next i
        
    If first Then
    difference = Closep - Openp
        ws.Cells(Ticker_Target.Row, outputcol).Value = difference
    Else
        ws.Cells(Ticker_Target.Row, outputcol).Value = "Not found"
    End If
    
Next Ticker_Target

'Third for loop to compare % changes for each variable
For Each Ticker_Target In Ticker_Targets
    first = False
    Openp = 0
    Closep = 0
    
        For i = 1 To lastrow
            If ws.Cells(i, "A").Value = Ticker_Target.Value Then
                If Not first Then
                    Openp = ws.Cells(i, "C").Value
                    first = True
                End If
                Closep = ws.Cells(i, "F").Value
            End If
        Next i
        
    If first Then
    percentage = (Closep - Openp) / Openp
        ws.Cells(Ticker_Target.Row, outputcol2).Value = percentage
    Else
        ws.Cells(Ticker_Target.Row, outputcol2).Value = "Not found"
    End If
    
Next Ticker_Target

'Formating %
ws.Columns("K:K").NumberFormat = "0.00%"
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

'Three separate for loops for the max and mins
maxpercent = -1E+300
minpercent = 1E+300

For i = 2 To lastrow
    If ws.Cells(i, 11).Value > maxpercent Then
        maxpercent = ws.Cells(i, 11).Value
        maxticker = ws.Cells(i, 9).Value
    End If
Next i

ws.Range("P2").Value = maxticker
ws.Range("Q2").Value = maxpercent

For i = 2 To lastrow
    If ws.Cells(i, 11).Value < minpercent Then
        minpercent = ws.Cells(i, 11).Value
        minticker = ws.Cells(i, 9).Value
    End If
Next i

ws.Range("P3").Value = minticker
ws.Range("Q3").Value = minpercent

For i = 2 To lastrow
    If ws.Cells(i, 12).Value > maxvol Then
        maxvol = ws.Cells(i, 12).Value
        maxvolticker = ws.Cells(i, 9).Value
    End If
Next i

ws.Range("P4").Value = maxvolticker
ws.Range("Q4").Value = maxvol

'For loop for color formatting
Dim TargetColColor As Integer
TargetColColor = 10

For i = 2 To lastrow
    If ws.Cells(i, TargetColColor).Value > 0 Then
        ws.Cells(i, TargetColColor).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, TargetColColor).Value < 0 Then
        ws.Cells(i, TargetColColor).Interior.ColorIndex = 3
    End If
Next i
   
Next ws

End Sub

