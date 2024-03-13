Sub Stocks()

Dim ws As Worksheet
Dim Tickername As String
Dim lastrow As Long
Dim Totalstockvolume As Double
Dim OpeningValue As Double
Dim ClosingValue As Double
Dim Flag As Boolean
Dim Flag2 As Boolean
Dim lastrow2 As Long
Dim yearlychange As Double
Dim percentchange As Double
Dim Greatestincrease As Double
Dim Greatestincreaseticker As String
Dim Greatestdecrease As Double
Dim Greatestdecreaseticker As String
Dim Greatesttotalvolume As Double
Dim Greatesttotalvolumeticker As String
Totalstockvolume = 0

Dim Summary_Table_Row As Integer
Dim Summary_Table_Row2 As Integer
Dim Summary_Table_Row3 As Integer

For Each ws In Sheets

Summary_Table_Row = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 14).Value = "Opening Value"
ws.Cells(1, 15).Value = "Closing Value"
ws.Range("R2").Value = "Greatest % Increase"
ws.Range("R3").Value = " Greatest % Decrease"
ws.Range("R4").Value = "Greatest Total Volume"
ws.Range("S1").Value = "Ticker"
ws.Range("T1").Value = "Value"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 Tickername = ws.Cells(i, 1).Value
 Totalstockvolume = Totalstockvolume + ws.Cells(i, 7).Value
  ws.Range("I" & Summary_Table_Row).Value = Tickername
  ws.Range("L" & Summary_Table_Row).Value = Totalstockvolume
  Summary_Table_Row = Summary_Table_Row + 1
  Totalstockvolume = 0
  Else
  Totalstockvolume = Totalstockvolume + ws.Cells(i, 7).Value
   End If
   Next i
   
  Summary_Table_Row2 = 2
  OpeningValue = 0
  Flag = False
  
For i = 2 To lastrow
 If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And (Flag = False) Then
 OpeningValue = ws.Cells(i, 3).Value
  ws.Range("N" & Summary_Table_Row2).Value = OpeningValue
  Flag = True
  ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  Flag = False
    Summary_Table_Row2 = Summary_Table_Row2 + 1
    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And (Flag = False) Then
 OpeningValue = ws.Cells(i, 3).Value
  ws.Range("N" & Summary_Table_Row2).Value = OpeningValue
  Flag = True
   End If
   Next i
 
 Flag2 = False
 Summary_Table_Row3 = 2
 ClosingValue = 0
 
 For i = 2 To lastrow
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And (Flag2 = False) Then
 ClosingValue = ws.Cells(i, 6).Value
  ws.Range("O" & Summary_Table_Row3).Value = ClosingValue
  Flag2 = True
  Summary_Table_Row3 = Summary_Table_Row3 + 1
  ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
  Flag2 = False
  ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And (Flag2 = False) Then
 ClosingValue = ws.Cells(i, 6).Value
  ws.Range("O" & Summary_Table_Row3).Value = ClosingValue
  Flag2 = True
   End If
   Next i
 
yearlychange = 0
 lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
 For j = 2 To lastrow2
 yearlychange = ws.Cells(j, 15).Value - ws.Cells(j, 14).Value
 ws.Cells(j, 10).Value = yearlychange
 Next j
 
  For j = 2 To lastrow2
If ws.Cells(j, 10).Value <= 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 3
Else
ws.Cells(j, 10).Interior.ColorIndex = 4
End If
 Next j
 
 percentchange = 0
  For j = 2 To lastrow2
  percentchange = (((ws.Cells(j, 14).Value - ws.Cells(j, 15).Value) / ws.Cells(j, 14).Value)) * -1
  ws.Cells(j, 11).Value = percentchange
  ws.Range("K:K").NumberFormat = "0.00%"
  Next j
  
    For j = 2 To lastrow2
If ws.Cells(j, 11).Value <= 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 3
Else
ws.Cells(j, 11).Interior.ColorIndex = 4
End If
 Next j
 
  Greatestincrease = 0
For j = 2 To lastrow2
If ws.Cells(j, 11).Value > Greatestincrease Then
Greatestincrease = ws.Cells(j, 11).Value
Greatestincreaseticker = ws.Cells(j, 9).Value
ws.Range("S2").Value = Greatestincreaseticker
ws.Range("T2").Value = Greatestincrease
End If
Next j

Greatestdecrease = 0
For j = 2 To lastrow2
If ws.Cells(j, 11).Value < Greatestdecrease Then
Greatestdecrease = ws.Cells(j, 11).Value
Greatestdecreaseticker = ws.Cells(j, 9).Value
ws.Range("S3").Value = Greatestdecreaseticker
ws.Range("T3").Value = Greatestdecrease
End If
Next j

ws.Range("T2:T3").NumberFormat = "0.00%"

Greatesttotalvolume = 0
For j = 2 To lastrow2
If ws.Cells(j, 12).Value > Greatesttotalvolume Then
Greatesttotalvolume = ws.Cells(j, 12).Value
Greatesttotalvolumeticker = ws.Cells(j, 9).Value
ws.Range("S4").Value = Greatesttotalvolumeticker
ws.Range("T4").Value = Greatesttotalvolume
End If
Next j

 
Next ws


End Sub


