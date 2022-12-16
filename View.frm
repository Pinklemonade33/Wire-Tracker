VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} View 
   Caption         =   "View"
   ClientHeight    =   3420
   ClientLeft      =   -2232
   ClientTop       =   -8148
   ClientWidth     =   3228
   OleObjectBlob   =   "View.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MultiPage1_Change()

End Sub


Private Sub UserForm_Click()

End Sub

Private Sub vcComboBox_Click()
Dim iValues() As Integer, a As Integer, answer As Boolean, check As Boolean
Dim iLowCuts() As Integer, iHighCuts() As Integer, iBulk() As Integer

vcInvList.Clear
vcInvTotal.Caption = ""

If vcLowCuts = True Then iLowCuts = invFindRangeValues(vcComboBox.list(vcComboBox.ListIndex), "LowCuts")
If vcHighCuts = True Then iHighCuts = invFindRangeValues(vcComboBox.list(vcComboBox.ListIndex), "HighCuts")
If vcBulk = True Then iBulk = invFindRangeValues(vcComboBox.list(vcComboBox.ListIndex), "Bulk")

Call FillArray(iValues, iLowCuts)
Call FillArray(iValues, iHighCuts)
Call FillArray(iValues, iBulk)

check = checkarray(iValues)
If check = True Then
 For a = 1 To UBound(iValues)
  If iValues(a) > 0 Then vcInvList.object.AddItem iValues(a)
  If vcInvTotal.Caption > "" Then vcInvTotal.Caption = vcInvTotal.Caption + iValues(a)
  If vcInvTotal.Caption = "" Then vcInvTotal.Caption = iValues(a)
 Next a
End If
If check = False And vcHighCuts = True And vcLowCuts = True And vcBulk = True Then
  MsgBox ("No inventory of the selected wire type was found.")
End If

End Sub

Private Sub vcComboBox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer, ExistingWire As Boolean
Set ws1 = ThisWorkbook.Worksheets("Saved")

ExistingWire = sCheckWire("Wire Name")
If ExistingWire = False Then MsgBox ("Selected wire does not exist")

Do Until vcComboBox.ListCount = 0
 vcComboBox.object.RemoveItem 0
Loop

Do
a = a + 1
 If ws1.Cells(a, 1).Value > "" Then
  b = b + 1
  If ws1.Cells(a, 1).Value = "Wire Name" Then Exit Do
  vcComboBox.object.AddItem ws1.Cells(a, 1).Value
 End If
Loop

End Sub

Private Sub vcComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
vcComboBox_Click
KeyCode = 0
End If
End Sub

Private Sub vcCountList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If vcCountList.object.ListCount > 0 Then
 On Error Resume Next
 vcCountTotal.Caption = vcCountTotal.Caption + -vcCountList.object.list(vcCountList.object.ListIndex)
 vcCountList.object.RemoveItem vcCountList.object.ListIndex
End If
End Sub


Private Sub vcCountList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim answer As String
If KeyCode = 8 Then
 answer = MsgBox("Are you sure you want to delete your count?", vbYesNo)
 If answer = vbYes Then
  vcCountList.Clear
  vcCountTotal.Caption = ""
 End If
 KeyCode = 0
End If
End Sub

Private Sub vcdiffButton_Click()
Dim iValues() As Integer, a As Integer, b As Integer
  
vcDiffList.Clear
vcDiffTotal.Caption = ""

If vcCountList.ListCount = 0 Then
 MsgBox ("Please enter your count")
 Exit Sub
End If
If vcComboBox.Value = "" Then
 MsgBox ("Please select a wire type")
 Exit Sub
End If

ReDim iValues(1 To vcCountList.object.ListCount)
For a = 1 To vcCountList.object.ListCount
 iValues(a) = vcCountList.object.list(b)
 b = b + 1
Next a

Call VerifyCount(iValues, vcComboBox.Value)

End Sub

Private Sub vcLowCuts_Click()
If vcComboBox.Value > "" Then vcComboBox_Click
End Sub

Private Sub vcHighCuts_Click()
If vcComboBox.Value > "" Then vcComboBox_Click
End Sub

Private Sub vcBulk_Click()
If vcComboBox.Value > "" Then vcComboBox_Click
End Sub

Private Sub vcOkay_Click()
Dim iValues() As Integer, a As Integer, b As Integer, iValues2() As Integer, answer As String
 
If vcCountList.ListCount = 0 Then
 MsgBox ("Please enter your count")
 Exit Sub
End If
If vcComboBox.Value = "" Then
 MsgBox ("Please select a wire type")
 Exit Sub
End If

Call vcdiffButton_Click

ReDim iValues(1 To vcCountList.object.ListCount)
For a = 1 To vcCountList.object.ListCount
 iValues(a) = vcCountList.object.list(b)
 b = b + 1
Next a
iValues2 = invFindAllRangeValues(vcComboBox.Value)

If vcDiffList.object.ListCount > 0 Then
 answer = MsgBox("Would you like to change the inventory to your count?", vbYesNo)
 If answer = vbYes Then
  Call SortRemove(iValues2, vcComboBox.Value)
  Call SortAdd(iValues, vcComboBox.Value)
  View.vcDiffList.Clear
  View.vcCountList.Clear
  View.vcInvList.Clear
  View.vcComboBox.Value = ""
  View.vcInvTotal.Caption = ""
  View.vcCountTotal.Caption = ""
  View.vcDiffTotal.Caption = ""
  MsgBox ("Count added to Inventory")
 End If
End If

If vcDiffList.object.ListCount = 0 Then
  View.vcDiffList.Clear
  View.vcCountList.Clear
  View.vcInvList.Clear
  View.vcComboBox.Value = ""
  View.vcInvTotal.Caption = ""
  View.vcCountTotal.Caption = ""
  View.vcDiffTotal.Caption = ""
End If

End Sub

Private Sub vcTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim a As Integer, b As Integer, FirstNum As Integer, LastNum As Integer, sum As Integer, OldSum, c As Integer, SumList() As Variant

If KeyCode = 13 Then
For a = 1 To Len(vcTextBox.object.Value)
 If Mid(vcTextBox.object.Value, a, 1) > " " Then
  FirstNum = a
  b = 1
  Do Until Mid(vcTextBox.object.Value, a + 1, 1) = " " Or Mid(vcTextBox.object.Value, a + 1, 1) = ""
   a = a + 1
   b = b + 1
  Loop
  LastNum = b
  vcCountList.object.AddItem Mid(vcTextBox.object.Value, FirstNum, LastNum)
  c = c + 1
  ReDim Preserve SumList(1 To c)
  SumList(c) = Mid(vcTextBox.object.Value, FirstNum, LastNum)
 End If
Next a
KeyCode = 0

 If vcCountTotal.Caption > "" Then
  OldSum = CStr(vcCountTotal.Caption)
  sum = OldSum
 End If
  
 For a = 1 To UBound(SumList)
  sum = sum + SumList(a)
 Next a
 vcCountTotal.Caption = sum
 vcTextBox.Value = ""
End If

End Sub

Private Sub vfAltCombo2_DropButtonClick()
If vfAltCombo.object.Value = "Changes" Then
 vfAltCombo.object.AddItem "Date"
 vfAltCombo.object.AddItem ""
If vfAltCombo.object.Value = "Sites" Then
End Sub

Private Sub vfComboBox_Click()
Dim iValues() As Integer, a As Integer, answer As Boolean, check As Boolean
Dim iLowCuts() As Integer, iHighCuts() As Integer, iBulk() As Integer

vfInvList.Clear
vfInvTotal.Caption = ""

If vfLowCuts = True Then iLowCuts = invFindRangeValues(vfComboBox.list(vfComboBox.ListIndex), "LowCuts")
If vfHighCuts = True Then iHighCuts = invFindRangeValues(vfComboBox.list(vfComboBox.ListIndex), "HighCuts")
If vfBulk = True Then iBulk = invFindRangeValues(vfComboBox.list(vfComboBox.ListIndex), "Bulk")

Call FillArray(iValues, iLowCuts)
Call FillArray(iValues, iHighCuts)
Call FillArray(iValues, iBulk)

check = checkarray(iValues)
If check = True Then
 For a = 1 To UBound(iValues)
  If iValues(a) > 0 Then vfInvList.object.AddItem iValues(a)
  If vfInvTotal.Caption > "" Then vfInvTotal.Caption = vfInvTotal.Caption + iValues(a)
  If vfInvTotal.Caption = "" Then vfInvTotal.Caption = iValues(a)
 Next a
End If
If check = False And vfHighCuts = True And vfLowCuts = True And vfBulk = True Then
  MsgBox ("No inventory of the selected wire type was found.")
End If

End Sub

Private Sub vfComboBox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer, ExistingWire As Boolean
Set ws1 = ThisWorkbook.Worksheets("Saved")

ExistingWire = sCheckWire("Wire Name")
If ExistingWire = False Then MsgBox ("Selected wire does not exist")

Do Until vfComboBox.ListCount = 0
 vfComboBox.object.RemoveItem 0
Loop

Do
a = a + 1
 If ws1.Cells(a, 1).Value > "" Then
  b = b + 1
  If ws1.Cells(a, 1).Value = "Wire Name" Then Exit Do
  vfComboBox.object.AddItem ws1.Cells(a, 1).Value
 End If
Loop

End Sub

Private Sub vfComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
End Sub

Private Sub vfInvList_Click()
If vfInvList.object.ListCount > 0 Then
 On Error Resume Next
 vfLengthBox.Value = vfInvList.object.list(vfInvList.object.ListIndex)
End If
End Sub

Private Sub vfLengthBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 Call vfPopulateInvList
End Sub

Private Sub vfLengthBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 13 Then
  Call vfPopulateInvList
  KeyCode = 0
 End If
End Sub

Private Sub vfmaxBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 Call vfPopulateInvList
End Sub

Private Sub vfmaxBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 13 Then
  Call vfPopulateInvList
  KeyCode = 0
 End If
End Sub

Private Sub vfLowCuts_Click()
Call vfPopulateInvList
End Sub

Private Sub vfHighCuts_Click()
Call vfPopulateInvList
End Sub

Private Sub vfBulk_Click()
Call vfPopulateInvList
End Sub

Private Sub vfSiteBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim iValues() As Variant, a As Integer
 If KeyCode = 13 Then
  iValues = SortLongIntoArray(vfSiteBox.Value)
  For a = 1 To UBound(iValues)
   vfSiteList.AddItem iValues(a)
  Next a
  KeyCode = 0
  vfSiteBox.Value = ""
 End If
End Sub

Private Sub vfSiteList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 If vfSiteList.ListCount > 0 Then
  On Error Resume Next
  vfSiteList.RemoveItem vfSiteList.ListIndex
 End If
End Sub
