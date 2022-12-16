VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveWire 
   Caption         =   "UserForm1"
   ClientHeight    =   900
   ClientLeft      =   -300
   ClientTop       =   -1176
   ClientWidth     =   300
   OleObjectBlob   =   "RemoveWire.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveWire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub rComboBox_Click()
Dim iValues() As Integer, a As Integer, answer As Boolean, check As Boolean
Dim iLowCuts() As Integer, iHighCuts() As Integer, iBulk() As Integer

rInvList.Clear
rInvTotal.Caption = ""

If rLowCuts = True Then iLowCuts = invFindRangeValues(rComboBox.list(rComboBox.ListIndex), "LowCuts")
If rHighCuts = True Then iHighCuts = invFindRangeValues(rComboBox.list(rComboBox.ListIndex), "HighCuts")
If rBulk = True Then iBulk = invFindRangeValues(rComboBox.list(rComboBox.ListIndex), "Bulk")

Call FillArray(iValues, iLowCuts)
Call FillArray(iValues, iHighCuts)
Call FillArray(iValues, iBulk)

check = checkarray(iValues)
If check = True Then
 For a = 1 To UBound(iValues)
  If iValues(a) > 0 Then rInvList.object.AddItem iValues(a)
  If rInvTotal.Caption > "" Then rInvTotal.Caption = rInvTotal.Caption + iValues(a)
  If rInvTotal.Caption = "" Then rInvTotal.Caption = iValues(a)
 Next a
End If
If check = False And rHighCuts = True And rLowCuts = True And rBulk = True Then
  MsgBox ("No inventory of the selected wire type was found.")
End If

End Sub

Private Sub rCombobox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer, ExistingWire As Boolean
Set ws1 = ThisWorkbook.Worksheets("Saved")

ExistingWire = sCheckWire("Wire Name")
If ExistingWire = False Then MsgBox ("Selected wire does not exist")

Do Until rComboBox.ListCount = 0
 rComboBox.object.RemoveItem 0
Loop

Do
a = a + 1
 If ws1.Cells(a, 1).Value > "" Then
  b = b + 1
  If ws1.Cells(a, 1).Value = "Wire Name" Then Exit Do
  rComboBox.object.AddItem ws1.Cells(a, 1).Value
 End If
Loop

End Sub

Private Sub rInvList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim sum As Integer
If rInvList.ListCount > 0 Then
 On Error Resume Next
 If rInvTotal.Caption = "" Then rInvTotal.Caption = 0
 If rRemoveTotal.Caption = "" Then rRemoveTotal.Caption = 0
 sum = rInvList.list(rInvList.ListIndex)
 rInvTotal.Caption = rInvTotal.Caption + -sum
 rRemoveTotal.Caption = rRemoveTotal.Caption + sum
 rRemoveList.AddItem rInvList.list(rInvList.ListIndex)
 rInvList.RemoveItem rInvList.ListIndex
End If
End Sub

Private Sub rRemoveList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim sum As Integer
If rRemoveList.ListCount > 0 Then
 On Error Resume Next
 If rInvTotal.Caption = "" Then rInvTotal.Caption = 0
 If rRemoveTotal.Caption = "" Then rRemoveTotal.Caption = 0
 sum = rRemoveList.list(rRemoveList.ListIndex)
 rInvTotal.Caption = rInvTotal.Caption + sum
 rRemoveTotal.Caption = rRemoveTotal.Caption + -sum
 rInvList.AddItem rRemoveList.list(rRemoveList.ListIndex)
 rRemoveList.RemoveItem rRemoveList.ListIndex
End If
End Sub

Private Sub rRemoveList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim answer As String

If KeyCode = 8 Then
 answer = MsgBox("Are you sure you want clear the list?", vbYesNo)
 If answer = vbYes Then rRemoveList.Clear
End If
 
End Sub

Private Sub rRemoveBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim iValues() As Variant, a As Integer, sum As Integer, iValues2() As Variant, iValues3() As Variant, b As Integer, diff As Integer
If rRemoveBox.Value = 0 Or rRemoveBox.Value = "" Then Exit Sub
If KeyCode = 13 Then
 iValues = SortLongIntoArray(rRemoveBox.Value)
 ReDim iValues2(1 To rInvList.ListCount)
 For a = 1 To rInvList.ListCount
  iValues2(a) = rInvList.list(a + -1)
 Next a
 If UBound(iValues) > UBound(iValues2) Then iValues3 = FindDiscrepancies(iValues, iValues2, 2)
 If UBound(iValues) < UBound(iValues2) Then iValues3 = FindDiscrepancies(iValues2, iValues, 2)
 For a = 1 To UBound(iValues)
  For b = 1 To rInvList.ListCount + -1
   If rInvList.list(b) = iValues3(a) Then
    diff = diff + rInvList.list(b)
    rInvList.RemoveItem b
    Exit For
   End If
  Next b
  rRemoveList.AddItem iValues3(a)
  sum = sum + iValues3(a)
 Next a
 If rRemoveTotal.Caption = "" Then rRemoveTotal.Caption = 0
  If rInvTotal.Caption = "" Then rInvTotal.Caption = 0
 rRemoveTotal.Caption = rRemoveTotal.Caption + sum
 rInvTotal.Caption = rInvTotal.Caption + -diff
 rRemoveBox.Value = ""
 KeyCode = 0
End If
End Sub

Private Sub rRemoveButton_Click()
Dim answer As String
If rRemoveList.ListCount = 0 Then Exit Sub
answer = MsgBox(rRemoveList.ListCount & " Items are being Removed from inventory. Are you sure you want to remove them?", vbYesNo)
If answer = vbYes Then
 Call SortRemove(rRemoveList.list, rComboBox.Value)
 rRemoveBox.Value = ""
 rRemoveList.Clear
 rComboBox.Value = ""
 rInvTotal.Caption = ""
 rRemoveTotal.Caption = ""
End If
End Sub

Private Sub UserForm_Click()

End Sub
