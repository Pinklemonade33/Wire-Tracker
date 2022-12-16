VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OldSettings 
   Caption         =   "Settings"
   ClientHeight    =   1860
   ClientLeft      =   -216
   ClientTop       =   -1272
   ClientWidth     =   3768
   OleObjectBlob   =   "OldSettings.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "OldSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sure As Boolean

Private Sub sBaseBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If sBaseBox.Value > "" Then
 If KeyCode = 13 Then
   sBaseList.AddItem sBaseBox.Value
   sBaseBox.Value = ""
 End If
End If
End Sub

Private Sub sBaseList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim answer As Boolean

If sure = False Then
answer = MsgBox("Are you sure you want to remove the selected cut?", vbYesNo)
End If
If sure = True Then answer = True

If answer = True Then
 sure = True
 sBaseList.object.RemoveItem sBaseList.object.ListIndex
End If
End Sub


Private Sub sBox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer
Set ws1 = ThisWorkbook.Worksheets("Saved")

Call sCheckWire("Wire Name")

Do Until sBox.ListCount = 0
 sBox.object.RemoveItem 0
Loop

Do
a = a + 1
 If ws1.Cells(a, 1).Value > "" Then
  b = b + 1
  If ws1.Cells(a, 1).Value = "Wire Name" Then Exit Do
  sBox.object.AddItem ws1.Cells(a, 1).Value
 End If
Loop

End Sub


Private Sub sBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
 Call EnterOrPop(sWireLabel, sBox.Value)
 KeyCode = 0
End If
End Sub

Private Sub sbox_click()
Call Populate(sBox.Value)
End Sub



Private Sub sClear_Click()
Dim obj As Object

For Each obj In Settings.Controls
 If TypeName(obj) = "TextBox" Then obj.Value = ""
 If TypeName(obj) = "ListBox" Then obj.Clear
 If TypeName(obj) = "ComboBox" Then obj.Value = ""
 If TypeName(obj) = "Label" Then
  If obj.Name = "sWireLabel" Then obj.Caption = ""
 End If
Next obj
  

End Sub

Private Sub sDelete_Click()
Call DeleteSettings(sWireLabel.Caption)
End Sub

Private Sub sMaxBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If sMaxBox.Value > "" Then
 If KeyCode = 13 Then
  If sSpecList.ListCount + -1 = sMaxList.ListCount Then
   sMaxList.AddItem sMaxBox.Value
   sMaxBox.Value = ""
   KeyCode = 0
   sSpecBox.SetFocus
   Else: MsgBox ("Please enter a specific cut")
  End If
 End If
End If
End Sub


Private Sub sMultiPage_Change()

End Sub

Private Sub sSpecBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If sSpecBox.Value > "" Then
 If KeyCode = 13 Then
  If sSpecList.ListCount = sThreshList.ListCount And sSpecList.ListCount = sMaxList.ListCount Then
   sSpecList.AddItem sSpecBox.Value
   sSpecBox.Value = ""
   Else: MsgBox ("Please set the thresholds for the previously entered wire")
  End If
 End If
End If
End Sub

Private Sub sSpecList_Click()
If sThreshList.ListCount = sSpecList.ListCount And sThreshList.ListCount = sSpecList.ListCount Then
 sThreshList.ListIndex = sSpecList.ListIndex
 sMaxList.ListIndex = sSpecList.ListIndex
End If
End Sub

Private Sub sSpecList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 Dim answer As Boolean

If sure = False Then
answer = MsgBox("Are you sure you want to remove the selected cut?", vbYesNo)
End If
If sure = True Then answer = True

If answer = True Then
 sure = True
 sSpecList.object.RemoveItem sSpecList.object.ListIndex
 sThreshList.object.RemoveItem sThreshList.object.ListIndex
 sMaxList.object.RemoveItem sMaxList.object.ListIndex
End If
End Sub

Private Sub sThreshBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If sThreshBox.Value > "" Then
 If KeyCode = 13 Then
  If sSpecList.ListCount + -1 = sThreshList.ListCount Then
   sThreshList.AddItem sThreshBox.Value
   sThreshBox.Value = ""
   Else: MsgBox ("Please enter a specific cut")
  End If
 End If
End If
End Sub

Private Sub wcExit_Click()
Settings.Hide
End Sub

Private Sub wcSave_Click()
If sWireLabel.Caption = "Label1" Then
 MsgBox ("Please enter wire name")
 Exit Sub
End If

Call SaveSettings(sWireLabel.Caption)
End Sub
