VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Settings"
   ClientHeight    =   2448
   ClientLeft      =   -1020
   ClientTop       =   -4116
   ClientWidth     =   4692
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sBaseBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If sBaseBox.Value > "" Then
 If KeyCode = 13 Then
   sBaseList.AddItem sBaseBox.Value
   sBaseBox.Value = ""
 End If
End If
End Sub

Private Sub sBaseList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim answer As String

answer = MsgBox("Are you sure you want to remove the selected cut?", vbYesNo)

If answer = vbYes Then
 sBaseList.object.RemoveItem sBaseList.object.ListIndex
End If

End Sub


Private Sub sBox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer, ExistingWire As Boolean
Set ws1 = ThisWorkbook.Worksheets("Saved")

ExistingWire = sCheckWire("Wire Name")
If ExistingWire = False Then MsgBox ("Selected wire does not exist")

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
ClearSettings
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
If sThreshList.ListCount = sSpecList.ListCount And sMaxList.ListCount = sSpecList.ListCount Then
 sThreshList.ListIndex = sSpecList.ListIndex
 sMaxList.ListIndex = sSpecList.ListIndex
End If
End Sub

Private Sub sSpecList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 Dim answer As String
 
answer = MsgBox("Are you sure you want to remove the selected cut?", vbYesNo)

If answer = vbYes Then
 If sSpecList.object.ListIndex + 1 <= sThreshList.object.ListCount Then sThreshList.object.RemoveItem sSpecList.object.ListIndex
 If sSpecList.object.ListIndex + 1 <= sMaxList.object.ListCount Then sMaxList.object.RemoveItem sSpecList.object.ListIndex
 sSpecList.object.RemoveItem sSpecList.object.ListIndex
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

Private Sub UserForm_initialize()
With Settings
 .Height = 600.6
 .Width = 1122
End With
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

