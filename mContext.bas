Attribute VB_Name = "mContext"
Option Explicit
Dim ContextMenu As CommandBar, ctrl As CommandBarControl, Pop As CommandBarPopup

Sub SetContextMenu()
Set ContextMenu = Application.CommandBars("cell")
End Sub

Sub test()
Dim ContextMenu As CommandBar, MySubMenu As CommandBarControl

Call DeleteFromCellMenu

Set ContextMenu = Application.CommandBars("Cell")

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
 .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleCaseMacro"
 .FaceId = 0
 .Caption = "Manual"
 .Tag = "My_Cell_Control_Tag"
End With


End Sub

Sub DeleteDefault()
Dim ContextMenu As CommandBar, ctrl As CommandBarControl
Application.ShowMenuFloaties = True
Set ContextMenu = Application.CommandBars("Cell")

For Each ctrl In ContextMenu.Controls
  ctrl.Delete
Next ctrl

End Sub

Sub DeleteCustom()
Dim ContextMenu As CommandBar, ctrl As CommandBarControl

Set ContextMenu = Application.CommandBars("Cell")

For Each ctrl In ContextMenu.Controls
 If ctrl.Tag = "My_Cell_Control_Tag" Then
  ctrl.Delete
 End If
Next ctrl

End Sub


Sub FindAllNames()
Dim cmdbar As CommandBar

For Each cmdbar In CommandBars
On Error Resume Next
 Debug.Print cmdbar.Name
Next cmdbar

End Sub

Sub FindStuff()
Dim ContextMenu As CommandBar, ctrl As CommandBarControl
Application.ShowMenuFloaties = True
Set ContextMenu = Application.CommandBars("Cell")
For Each ctrl In ContextMenu.Controls
 Debug.Print ctrl.Caption
Next ctrl

End Sub

Sub dostuff()
Dim ContextMenu As CommandBar, ctrl As CommandBarControl

Set ContextMenu = Application.CommandBars("Cell")

For Each ctrl In ContextMenu.Controls
 If ctrl.Caption = "S&ort" Then Debug.Print ctrl.Type
Next ctrl

End Sub

Sub AddWireTypeContext()
Call SetContextMenu
Call DeleteCustom
Call DeleteDefault
With ContextMenu.Controls.Add(Type:=10, before:=1)
 .Caption = "Manual"
 .Tag = "My_Cell_Control_Tag"
End With

Set Pop = ContextMenu.Controls(1)
With Pop.Controls.Add(Type:=1)
 .Caption = "Start Manual on Cut"
 .Tag = "My_Cell_Control_Tag"
End With
With Pop.Controls.Add(Type:=1)
 .Caption = "Start Manual on Trim"
 .Tag = "My_Cell_Control_Tag"
End With
With Pop.Controls.Add(Type:=1)
 .Caption = "Start Manual on Calculation"
 .Tag = "My_Cell_Control_Tag"
End With
With Pop.Controls.Add(Type:=1)
 .Caption = "Start Manual on Start"
 .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=1, before:=2)
 .OnAction = "'" & ThisWorkbook.Name & "'!" & "ChangeSettings"
 .Caption = "Change Settings"
 .Tag = "My_Cell_Control_Tag"
End With

End Sub

Sub RestoreDefault()
Call SetContextMenu
Call DeleteCustom
ContextMenu.Reset
Application.ShowMenuFloaties = False
End Sub

Sub ChangeSettings()
Settings.sBox.Value = ActiveCell.Value
Call Populate(ActiveCell.Value)
Settings.Show
End Sub

Sub ChangeManualType(cell, iType)

End Sub
