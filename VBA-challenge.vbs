Attribute VB_Name = "Module1"
Sub Vba_Challenge()

'Create loop through all sheets
For Each ws In Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name

'create variables to hold in summary table 1
Dim Ticker_Symbol As String
Dim Volume_Total As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Ticker_Summary_Row As Integer

'create variables to hold in summary tablke 2
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Long


'set inital values
Yearly_Change = 0
Volume_Total = 0
Percentage_Change = 0
Ticker_Summary_Row = 2

'Cell Labels & Format for summary table 1
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("J:J").NumberFormat = "0.00"
ws.Range("K:K").NumberFormat = "0.00%"

'create lastrow that keeps track of the last row through all worksheets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set k inital open value
k = 2

'create for loop that loops through all ticker symbol/values
    For i = 2 To lastrow

'check to see if we are still in the same ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
    
    'Calulate Yealy Change
            Yearly_Change = ws.Cells(i, 6) - ws.Cells(k, 3)

    'Calulate Percentage Change
            Percent_Change = Yearly_Change / ws.Cells(k, 3)
            
    'Add to Total Stock Volume
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value

    'print ticker symbols in ticker summary row
            ws.Range("I" & Ticker_Summary_Row).Value = Ticker_Symbol

    'print yearly change in ticker summary row
            ws.Range("J" & Ticker_Summary_Row).Value = Yearly_Change

    'print percent change in ticker summary row
            ws.Range("K" & Ticker_Summary_Row).Value = Percent_Change

    'print the ticker volume in the ticker summary row
            ws.Range("L" & Ticker_Summary_Row).Value = Volume_Total

    'Add one to ticker summary row
            Ticker_Summary_Row = Ticker_Summary_Row + 1

    'reset total stock volume
            Volume_Total = 0

    'reset yearly change
            Yearly_Change = 0

    'reset percent change
            Percent_Change = 0

    're-define k value to calulate new open value
            k = i + 1

    'if cell immediately following row is the same ticker symbol
        Else

    'Add to to the Volume Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value


        End If
        
    Next i

'condtional formatting for yearly change

For i = 2 To lastrow

        If ws.Cells(i, 10).Value > 0 Then

            ws.Cells(i, 10).Interior.ColorIndex = 4
      
        ElseIf ws.Cells(i, 10).Value < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        Else: ws.Cells(i, 10).Interior.ColorIndex = white
        
        End If
    
    Next i


'Cell Labels & Format for summary table 2
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"


'Grestest Increase, Decrease and Volume
'Find Max and Min and print values in summary table 2
    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
    ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
    
'Match tickers to the corresponding greatest percent increase, percent decrease and total volume values
    Greatest_Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    Greatest_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    Greatest_Volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

'Print values in summary table 2
    ws.Range("P2") = ws.Cells(Greatest_Increase + 1, 9)
    ws.Range("P3") = ws.Cells(Greatest_Decrease + 1, 9)
    ws.Range("P4") = ws.Cells(Greatest_Volume + 1, 9)


        
    Next ws

End Sub
