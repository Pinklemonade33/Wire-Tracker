Attribute VB_Name = "mAddObjects"
Option Explicit

Sub SetButtonsForCalc()
Dim ca As Worksheet
Set ca = ThisWorkbook.Worksheets("Calculate")
Call AddButton(2, 1, 2, "Select2", "Select", "Calculate")
Call AddButton(4, 1, 2, "select3", "Select", "Calculate")
Call AddButton(6, 1, 2, "StartCalculation", "Start", "Calculate")
Call AddButton(6, 3, 4, "RefreshList", "Refresh List", "Calculate")
Call AddButton(6, 5, 5, "Clear", "Clear", "Calculate")
Call AddButton(2, 6, 6, "CommandButton1", "On Start", "Calculate")
Call AddButton(3, 6, 6, "CommandButton2", "On Cut", "Calculate")
Call AddButton(4, 6, 6, "CommandButton3", "On Trim", "Calculate")
Call AddButton(5, 6, 6, "CommandButton4", "On Calc", "Calculate")
Call AddButton(6, 6, 7, "CreateGroup", "Create Group", "Calculate")
Call AddButton(6, 8, 9, "DeleteGroup", "Delete Group", "Calculate")

End Sub


Sub AddButton(inputrow, inputfirstcol, inputlastcol, inputname, InputCaption, inputworksheet)
Dim btn As Object, r As Range

For Each btn In ThisWorkbook.Worksheets(inputworksheet).OLEObjects
If btn.Name = inputname Then ThisWorkbook.Worksheets(inputworksheet).OLEObjects(inputname).Delete
Next btn

Set r = Range(Worksheets(inputworksheet).Cells(inputrow, inputfirstcol), Worksheets(inputworksheet).Cells(inputrow, inputlastcol))
Set btn = Worksheets(inputworksheet).OLEObjects.Add(ClassType:="Forms.CommandButton.1", Left:=r.Left, top:=r.top, Width:=r.Width, Height:=r.Height)
With btn
 .Name = inputname
 With .object
  .Caption = InputCaption
  With .Font
    .Size = 13
    .Bold = False
  End With
 End With
End With

End Sub

Sub AddComboBox(inputrow, inputfirstcol, inputlastcol, inputname, inputworksheet)
Dim Cbox As Object, r As Range

For Each Cbox In ThisWorkbook.Worksheets(inputworksheet).OLEObjects
 If Cbox.Name = inputname Then ThisWorkbook.Worksheets(inputworksheet).OLEObjects(inputname).Delete
Next Cbox

Set r = Range(Worksheets(inputworksheet).Cells(inputrow, inputfirstcol), Worksheets(inputworksheet).Cells(inputrow, inputlastcol))
Set Cbox = Worksheets(inputworksheet).OLEObjects.Add(ClassType:="Forms.ComboBox.1", Left:=r.Left, top:=r.top, Width:=r.Width, Height:=r.Height)

End Sub

Sub AddSpinControl(inputrow, inputfirstcol, inputlastcol, inputname, inputworksheet)
Dim Spin As Object, r As Range

For Each Spin In ThisWorkbook.Worksheets(inputworksheet).OLEObjects
 If Spin.Name = inputname Then ThisWorkbook.Worksheets(inputworksheet).OLEObjects(inputname).Delete
Next Spin

Set r = Range(Worksheets(inputworksheet).Cells(inputrow, inputfirstcol), Worksheets(inputworksheet).Cells(inputrow, inputlastcol))
Set Spin = Worksheets(inputworksheet).OLEObjects.Add(ClassType:="Forms.SpinButton.1", Left:=r.Left, top:=r.top, Width:=r.Width, Height:=r.Height)
Spin.object.Orientation = 0

End Sub

