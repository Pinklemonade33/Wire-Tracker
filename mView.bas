Attribute VB_Name = "mView"
Option Explicit
Dim DateRange As Range

Sub vfPopulateList()
Dim iDates() As Date, iValues() As Integer, iSites() As Variant, ws As Worksheet, a As Integer
Dim arr As Variant, sum As Integer, cell As Range
Set ws = ThisWorkbook.Worksheets("ViewTemp")

View.vfSiteList.Clear
View.vfAltList.Clear
View.vfAltTotal.Caption = ""
ws.Cells.Value = ""

iSites = pCollectFromSites(View.vfComboBox.Value, View.vfSiteList.list, View.vfAltCombo.Value, _
View.vfMaxBox.Value, View.vfLengthBox.Value, View.vfDateBeforeBox.Value, View.vfDateAfterBox.Value, "Sites")

iValues = pCollectFromSites(View.vfComboBox.Value, View.vfSiteList.list, View.vfAltCombo.Value, _
View.vfMaxBox.Value, View.vfLengthBox.Value, View.vfDateBeforeBox.Value, View.vfDateAfterBox.Value, "Values")

iDates = pCollectFromSites(View.vfComboBox.Value, View.vfSiteList.list, View.vfAltCombo.Value, _
View.vfMaxBox.Value, View.vfLengthBox.Value, View.vfDateBeforeBox.Value, View.vfDateAfterBox.Value, "Dates")

Set DateRange = Range(ws.Cells(1, 1), ws.Cells(UBound(iDates), 1))
For Each cell In DateRange
 a = a + 1
 cell.Value = iDates(a)
Next cell

For Each arr In iSites
 View.vfSiteList.AddItem arr
Next arr

For Each arr In iValues
 View.vfAltList.AddItem arr
 sum = sum + arr
Next arr
View.vfAltTotal.Caption = sum

End Sub

Sub vfPopulateInvList()
Dim a As Integer, check As Boolean, iLowCuts() As Integer, iHighCuts() As Integer, iBulk() As Integer
Dim iValues() As Integer
If View.vfComboBox.Value = "" Then Exit Sub
View.vfInvList.Clear
View.vfInvTotal.Caption = ""

If View.vfLowCuts = True Then iLowCuts = invFindRangeValues(View.vfComboBox.list(View.vfComboBox.ListIndex), "LowCuts")
If View.vfHighCuts = True Then iHighCuts = invFindRangeValues(View.vfComboBox.list(View.vfComboBox.ListIndex), "HighCuts")
If View.vfBulk = True Then iBulk = invFindRangeValues(View.vfComboBox.list(View.vfComboBox.ListIndex), "Bulk")

Call FillArray(iValues, iLowCuts)
Call FillArray(iValues, iHighCuts)
Call FillArray(iValues, iBulk)

check = checkarray(iValues)
If check = True Then
 For a = 1 To UBound(iValues)
  If iValues(a) > 0 Then
   If iValues(a) >= View.vfLengthBox.Value And iValues(a) <= View.vfMaxBox.Value And View.vfMaxBox.Value > "" And View.vfLengthBox.Value > "" _
   Or iValues(a) = View.vfLengthBox.Value And View.vfMaxBox.Value = "" And View.vfLengthBox.Value > "" _
   Or View.vfMaxBox.Value = "" And View.vfLengthBox.Value = "" Then View.vfInvList.AddItem iValues(a)
  End If
  
  If View.vfInvTotal.Caption > "" Then View.vfInvTotal.Caption = View.vfInvTotal.Caption + iValues(a)
  If View.vfInvTotal.Caption = "" Then View.vfInvTotal.Caption = iValues(a)
 Next a
End If
If check = False And View.vfHighCuts = True And View.vfLowCuts = True And View.vfBulk = True Then
  MsgBox ("No inventory of the selected wire type was found.")
End If

End Sub
