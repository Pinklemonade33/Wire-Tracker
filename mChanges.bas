Attribute VB_Name = "mChanges"
Option Explicit

Sub cAddToChanges(mType, mValues, wirename)
Dim FirstRow As Integer, iValues() As Variant, ws As Worksheet, LastRow As Integer
Set ws = ThisWorkbook.Worksheets("Changes")

FirstRow = FindNewRow
LastRow = FirstRow + UBound(mValues) + -1

Range(ws.Cells(FirstRow, 1), ws.Cells(LastRow, 1)).Value = wirename
Range(ws.Cells(FirstRow, 2), ws.Cells(LastRow, 2)).Value = mValues
Range(ws.Cells(FirstRow, 3), ws.Cells(LastRow, 3)).Value = mType
Range(ws.Cells(FirstRow, 4), ws.Cells(LastRow, 4)).Value = Now

End Sub

Function FindNewRow()
Dim a As Integer, ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Changes")

Do Until ws.Cells(a + 1, 1).Value = ""
a = a + 1
Loop

FindNewRow = a

End Function

Function cFindRanges(mType, mValue, wirename, mDate)
Dim a As Integer, ws As Worksheet, Matched As Boolean, oValues() As Integer, b As Integer

Do Until ws.Cells(a + 1, 1).Value = ""
a = a + 1
Matched = CheckRange(mType, mValue, wirename, mDate, a)
 If Matched = True Then
  b = b + 1
  ReDim Preserve oValues(1 To b)
  oValues(b) = a
 End If
Loop

End Function

Function CheckRange(mType, mValue, wirename, mDate, mRow)
Dim check(1 To 4) As Boolean, ws As Worksheet, a As Integer
Set ws = ThisWorkbook.Worksheets("Changes")

If wirename = False Then check(1) = True
If mValue = False Then check(2) = True
If mType = False Then check(3) = True
If mDate = False Then check(4) = True

If check(1) = False Then
 If ws.Cells(mRow, 1).Value = wirename Then check(1) = True
End If
If check(2) = False Then
 If ws.Cells(mRow, 2).Value >= mValue(1) _
 And ws.Cells(mRow, 2).Value <= mValue(2) Then check(2) = True
End If
If check(3) = False Then
 If ws.Cells(mRow, 3).Value = mType Then check(3) = True
End If
If check(4) = False Then
 If ws.Cells(mRow, 4).Value >= mDate(1) _
 And ws.Cells(mRow, 4).Value <= mDate(2) Then check(4) = True
End If

For a = 1 To 4
 If check(a) = False Then CheckRange = False
Next a

End Function

Function GetRangeValue(mRow, mCol)
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Changes")

If mCol = 1 Then GetRangeValue = ws.Cells(mRow, 1).Value
If mCol = 2 Then GetRangeValue = ws.Cells(mRow, 2).Value
If mCol = 3 Then GetRangeValue = ws.Cells(mRow, 3).Value
If mCol = 4 Then GetRangeValue = ws.Cells(mRow, 4).Value

End Function

Function cGetAllRangeValues(mRows, mCol)
Dim a As Integer, oValues() As Variant
ReDim oValues(1 To UBound(mRows))

For a = 1 To UBound(mRows)
 oValues(a) = GetRangeValue(mRows(a), mCol)
Next a

GetAllRangeValues = oValues

End Function
