Attribute VB_Name = "Module1"
Sub hw2():



Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
vol = 0
Dim rowspast As Double
rowspast = 0
'this prevents my overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table_Row = 2

    'loop
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        vol = vol + Cells(i, 7).Value
        
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'find all the values
            ticker = ws.Cells(i, 1).Value
            
            year_open = ws.Cells(i - rowspast, 3).Value
            year_close = ws.Cells(i, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close

            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            Select Case percent_change
               Case Is >= 0
                 ws.Cells(Summary_Table_Row, 11).Interior.Color = vbGreen
                 
               Case Is < 0
                  ws.Cells(Summary_Table_Row, 11).Interior.Color = vbRed
                 
            End Select
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

             vol = 0
             rowspast = -1
        End If
rowspast = rowspast + 1
'finish loop
    Next i
    
ws.Columns("K").NumberFormat = "0.00%"

'move to next worksheet
Next ws

End Sub
