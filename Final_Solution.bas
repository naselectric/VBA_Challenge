Attribute VB_Name = "Final_Solution"
Sub home2()


WS_Count = ActiveWorkbook.Worksheets.Count

         
For n = 1 To WS_Count

ActiveWorkbook.Worksheets(n).Select


a = ActiveSheet.UsedRange.Rows.Count
b = a - 1

Cells(2, 9).Value = Cells(2, 1)
Cells(2, 10).Value = Cells(2, 3)

j = 1
totalstock = 0

For I = 1 To b


If Cells(2 + I, 1).Value <> Cells(1 + I, 1).Value Then


j = 1 + j

Cells(j, 11).Value = Cells(1 + I, 6)
Cells(1 + j, 9).Value = Cells(2 + I, 1)
Cells(1 + j, 10).Value = Cells(2 + I, 3)
Cells(j, 14).Value = totalstock
totalstock = 0

End If

totalstock = Cells(2 + I, 7).Value + totalstock

Next I

Range("L2").Formula = "=K2-J2"
Range("M2").Formula = "=(L2/J2)"
Range("L2:M2").Select
Selection.Copy
Range("L3:" & "M" & j).PasteSpecial
Range("M2:" & "M" & j).NumberFormat = "0%"

Next n

MsgBox ("hi")

End Sub
