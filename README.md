# hw2-VBA
VBA Homework

Sub StockData2():

Dim ws As Worksheet
Dim Ticker As String
Dim Volume As Long
Dim TSV As Double
Dim j As Integer



For Each ws In Worksheets
  j = 0
  TSV = 0

LastRow = Cells(Rows.Count, "A").End(xlUp).Row
       For i = 2 To LastRow



    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


        TSV = ws.Cells(i, 7).Value + TSV


       ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = TSV

        TSV = 0
j = j + 1
  Else
      TSV = TSV + ws.Cells(i, 7).Value
        End If

    Next i

TSV = 0
j = 0

Next ws
End Sub
