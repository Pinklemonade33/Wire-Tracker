Attribute VB_Name = "mHome"
Option Explicit
Dim dWireList As Range

Sub UpdateChanges()
Dim cRanges() As Integer, a As Integer, cNow(1 To 2) As Date, CurrentValue As Variant, ws As Worksheet
Dim iDates() As Date, iValues() As Integer, iWireType() As Variant, iType As String, col As Integer
Set ws = ThisWorkbook.Worksheets("Home")
col = FindObjectCell("dWireLabel", ws)
Set dWireList = Range(ws.Cells(2, 1), ws.Cells(24, col))

cNow(1) = Now
cNow(2) = Now + ("00,00,14")


cRanges = cFindRanges(False, False, False, cNow)
iWireType = cGetAllRangeValues(cRanges, 1)
iValues = cGetAllRangeValues(cRanges, 2)
iType = cGetAllRangeValues(cRanges, 3)
iDates = cGetAllRangeValues(cRanges, 4)

dWireList.Value = ""
dWireList.Font.ColorIndex = 1

For a = 1 To UBound(cRanges)
 dWireList.Cells(a + 1, col).Value = iWireType(a) & " - " & iValues(a) & " - " & iType(a) & " - " & iDates(a)
 Call ColorFont(dWireList, iType(a), a)
Next a

End Sub

Sub ColorFont(mRange, mType, mRow)
If mType = "Added" Then mRange.Cells(a, 1).Font.Color = RGB(0, 176, 80)
If mType = "Removed" Then mRange.Cells(a, 1).Font.Color = RGB(255, 0, 0)
If mType = "Picked" Then mRange.Cells(a, 1).Font.Color = RGB(255, 0, 0)
If mType = "UnPicked" Then mRange.Cells(a, 1).Font.Color = RGB(0, 176, 80)
End Sub

Sub FindObjectCell(obj, ws)
Dim cell As Range

For Each cell In Range(ws.Cells(1, 1), ws.Cells(25, 25))
If cell.Left = ActiveSheet.OLEObjects(obj).Left Then
 FindObjectCell = cell.Row
 Exit For
End If
Next cell

End Sub


Sub UpdateLowWire()
Dim ws As Worksheet, LowWire As Range, FirstRow As Integer, col As Integer
Set ws = ThisWorkbook.Worksheets("Home")

col = FindObjectCell("dLowWireLabel", ws)
Set LowWire = Range(ws.Cells(2, col), ws.Cells(24, col))
LowWire.Value = iCollectLowWire


End Sub
