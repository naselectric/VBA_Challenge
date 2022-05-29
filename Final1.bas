Attribute VB_Name = "Final1"
Sub home2()


WS_Count = ActiveWorkbook.Worksheets.Count

         
For n = 1 To WS_Count

ActiveWorkbook.Worksheets(n).Select


Columns("I:P").AutoFit
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Stock Volume"
Range("N2") = "Greatest % Increase"
Range("N3") = "Greatest % Decrease"
Range("N4") = "Greatest Total Volume"
Range("O1") = "Ticker"
Range("P1") = "Value"



Range("I2").Value = Cells(2, 1)
x1 = Range("C2")
totalstock = 0


a = ActiveSheet.UsedRange.Rows.Count
b = a - 1

j = 1

GPI = 0
GPD = 1000000
GTV = 0

For I = 1 To b

totalstock = Cells(1 + I, 7).Value + totalstock

If Cells(2 + I, 1).Value <> Cells(1 + I, 1).Value Then

j = j + 1


y1 = Cells(1 + I, 6)

Cells(j, 10).Value = y1 - x1

z1 = y1 - x1

If z1 <= 0 Then

Cells(j, 10).Interior.Color = vbRed

Else

Cells(j, 10).Interior.Color = vbGreen

End If

z2 = z1 / (x1 + 0.000001)

Cells(j, 11).Value = z2

If z2 > GPI Then

GPI = z2
Ticker1 = Cells(j, 9).Value

End If

If z2 < GPD Then

GPD = z2
Ticker2 = Cells(j, 9).Value

End If

Cells(j, 12).Value = totalstock

If totalstock > GTV Then

GTV = totalstock
Ticker3 = Cells(j, 9).Value

End If

Cells(1 + j, 9).Value = Cells(2 + I, 1)
x1 = Cells(2 + I, 3).Value
totalstock = 0


End If


Next I

Range("K2:" & "K" & j).NumberFormat = "0%"
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"
Range("P4").NumberFormat = "0.0000E+0"
Range("P2").Value = GPI
Range("O2").Value = Ticker1
Range("P3").Value = GPD
Range("O3").Value = Ticker2
Range("P4").Value = GTV
Range("O4").Value = Ticker3

Next n

MsgBox ("hi " & b)

End Sub

