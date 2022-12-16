Attribute VB_Name = "mSettings"
Option Explicit
Dim Mode As Range, ManualStart As Range, CalcStart As Range, PrefTrim As Range, grouping As Range, ws1 As Worksheet, Orders As Range
Dim BaseInc As Range, Thresh As Range, MaxThresh As Range, WireNameRange As Range, NewRange As Range, All As Range, SpecCuts As Range

Function sCheckWire(wirename)
Dim a As Integer, ExistingWire As Boolean
Set ws1 = ThisWorkbook.Worksheets("Saved")

If IsNumeric(wirename) Then wirename = CDbl(wirename)

If ws1.Cells(1, 1).Value = "" Then
 MsgBox ("Saved Sheet is not aligned properly")
 End
End If

Do
a = a + 1
If ws1.Cells(a, 1).Value = wirename Then
 ExistingWire = True
 Exit Do
End If
If ws1.Cells(a, 1).Value = "Wire Name" Then
 ExistingWire = False
 Exit Do
End If
Loop

sCheckWire = ExistingWire

End Function

Sub FindSaveRanges(wirename)
Dim cell As Range, ListB As Object, HighestList As Integer
Dim a As Integer, LastRow As Integer, ExistingWire As Boolean
Dim SpecRows As Integer, BaseRows As Integer

ExistingWire = sCheckWire(wirename)

Do
a = a + 1
If ws1.Cells(a, 1).Value = wirename Then
 Set WireNameRange = ws1.Cells(a, 1)
 Exit Do
End If
Loop

If ExistingWire = True Then
 Do
  a = a + 1
  If ws1.Cells(a, 3).Value > "" Then BaseRows = BaseRows + 1
  If ws1.Cells(a, 4).Value > "" Then SpecRows = SpecRows + 1
  If ws1.Cells(a, 3).Value = "" And ws1.Cells(a, 4).Value = "" Then Exit Do
 Loop
Else
BaseRows = Settings.sBaseList.object.ListCount
SpecRows = Settings.sSpecList.object.ListCount
End If

Call SetRanges(WireNameRange, Mode, 3, 2)
Call SetRanges(Mode, ManualStart, 4, 2)
Call SetRanges(ManualStart, CalcStart, 1, 2)
Call SetRanges(CalcStart, PrefTrim, 1, 2)
Call SetRanges(PrefTrim, grouping, 1, 2)
Call SetRanges(grouping, Orders, 1, 2)

Call SetRanges(WireNameRange, BaseInc, BaseRows, 3)
Call SetRanges(WireNameRange, SpecCuts, SpecRows, 4)
Call SetRanges(WireNameRange, Thresh, SpecRows, 5)
Call SetRanges(WireNameRange, MaxThresh, SpecRows, 6)

HighestList = 11
If BaseRows > HighestList Then HighestList = BaseRows
If SpecRows > HighestList Then HighestList = SpecRows

LastRow = HighestList

Set NewRange = ws1.Cells(WireNameRange.Rows(1).Row + HighestList + 1, 1)
Set All = Range(ws1.Cells(WireNameRange.Rows(1).Row, 1), ws1.Cells(WireNameRange.Rows(1).Row + HighestList, 6))

End Sub

Sub SaveSettings(wirename)
Dim ExistingWire As Boolean
ExistingWire = sCheckWire(wirename)

If ExistingWire = True Then
 Call FindSaveRanges(wirename)
 All.Delete (xlShiftUp)
 ExistingWire = sCheckWire(wirename)
End If

Call FindSaveRanges("Wire Name")

With Settings
 WireNameRange.Value = .sWireLabel.Caption
 If .smStandard = True Then Mode.Rows(1).Value = .smStandard.Caption
 If .smCritical = True Then Mode.Rows(2).Value = .smCritical.Caption
 If .smFree = True Then Mode.Rows(3).Value = .smFree.Caption
 If .slmStart = True Then ManualStart.Rows(1).Value = .slmStart.Caption
 If .slmCalc = True Then ManualStart.Rows(2).Value = .slmCalc.Caption
 If .slmCut = True Then ManualStart.Rows(3).Value = .slmCut.Caption
 If .slmTrim = True Then ManualStart.Rows(4).Value = .slmTrim.Caption
 CalcStart.Value = .sCalcBox.Value
 PrefTrim.Value = .sTrimBox.Value
 grouping.Value = .sGroupingBox
 Orders.Value = .sOrdersBox
 BaseInc.Value = .sBaseList.list
 SpecCuts.Value = .sSpecList.list
 Thresh.Value = .sThreshList.list
 MaxThresh.Value = .sThreshList.list
End With
NewRange.Value = "Wire Name"

MsgBox (Settings.sWireLabel.Caption & " Saved")
    
End Sub

Sub DeleteSettings(wirename)
Dim answer As Boolean, ExistingWire As Boolean

ExistingWire = sCheckWire(wirename)

If ExistingWire = False Then
 MsgBox ("Selected Wire Not Found")
 Exit Sub
End If

Call FindSaveRanges(wirename)

answer = MsgBox("Are you sure you want to delete " & Settings.sWireLabel.Caption & "?", vbYesNo)
If answer = True Then
 All.Delete (xlShiftUp)
 MsgBox (Settings.sWireLabel & " Deleted")
End If

End Sub


Sub SetRanges(FirstRange, SecondRange, NumberOfRows, col)
Dim FirstRow As Integer

FirstRow = FirstRange.Rows(1).Row + FirstRange.Rows.Count
Set SecondRange = Range(ws1.Cells(FirstRow, col), ws1.Cells(FirstRow + NumberOfRows + -1, col))

End Sub

Sub Populate(wirename)
Dim ExistingWire As Boolean
ExistingWire = sCheckWire(wirename)

ExistingWire = sCheckWire(wirename)
If ExistingWire = False Then
 MsgBox ("Selected Wire Not Found")
 Exit Sub
End If

Call FindSaveRanges(wirename)
Call ClearSettings

With Settings
 .sWireLabel.Caption = WireNameRange.Value
 If Mode.Rows(1).Value = .smStandard.Caption Then .smStandard = True
 If Mode.Rows(2).Value = .smCritical.Caption Then .smCritical = True
 If Mode.Rows(3).Value = .smFree.Caption Then .smFree = True
 If ManualStart.Rows(1).Value = .slmStart.Caption Then .slmStart = True
 If ManualStart.Rows(2).Value = .slmCalc.Caption Then .slmCalc = True
 If ManualStart.Rows(3).Value = .slmCut.Caption Then .slmCut = True
 If ManualStart.Rows(4).Value = .slmTrim.Caption Then .slmTrim = True
 .sCalcBox.Value = CalcStart.Value
 .sTrimBox.Value = PrefTrim.Value
 .sGroupingBox = grouping.Value
 .sOrdersBox = Orders.Value
 If BaseInc.Rows.Count > 1 Then
  .sBaseList.list = BaseInc.Value
  Else: .sBaseList.AddItem BaseInc.Value
 End If
 If SpecCuts.Rows.Count > 1 Then
  .sSpecList.list = SpecCuts.Value
  Else: .sSpecList.AddItem SpecCuts.Value
 End If
 If SpecCuts.Rows.Count > 1 Then
  .sThreshList.list = Thresh.Value
  Else: .sThreshList.AddItem Thresh.Value
 End If
 If SpecCuts.Rows.Count > 1 Then
 .sMaxList.list = MaxThresh.Value
 Else: .sMaxList.AddItem MaxThresh.Value
 End If
 .sWireLabel.Caption = WireNameRange.Value
End With

End Sub

Sub EnterOrPop(objName, wirename)
Dim ExistingWire As Boolean

ExistingWire = sCheckWire(wirename)

If ExistingWire = False Then
  objName.Caption = wirename
End If
If ExistingWire = True Then
 Call Populate(wirename)
End If

End Sub

Sub ClearSettings()
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

Function FindCalcStart(wirename)
FindCalcStart = CalcStart.Value
End Function

Function FindTrim(wirename)
FindTrim = PrefTrim.Value
End Function

Function FindGrouping(wirename)
FindGrouping = grouping.Value
End Function

Function FindMode(wirename)
Dim cell As Range, oValues As Variant
For Each cell In Mode
 If cell.Value > "" Then
  oValues = cell.Value
  FindMode = oValues
 End If
Next cell
End Function

Function FindOrders(wirename)
FindOrders = Order.Value
End Function

Function FindManual(wirename)

Dim oValues() As Variant, cell As Range, a As Integer
For Each cell In ManualStart
 If cell.Value > "" Then
  a = a + 1
  ReDim Preserve oValues(1 To a)
  oValues(a) = cell.Value
 End If
Next cell

FindManual = oValues
End Function

Function FindBase(wirename)
Dim oValues() As Integer
oValues = CollectValues(BaseInc)
FindBase = oValues
End Function

Function FindSpec(wirename)
Dim oValues() As Integer
oValues = CollectValues(SpecCuts)
FindSpec = oValues
End Function

Function FindThresh(wirename)
Dim oValues() As Integer
oValues = CollectValues(Thresh)
FindThresh = oValues
End Function

Function FindMaxThresh(wirename)
Dim oValues() As Integer
oValues = CollectValues(MaxThresh)
FindMaxThresh = oValues
End Function

Function CollectValues(mRange)
Dim cell As Range, oValues() As Integer, a As Integer

For Each cell In mRange
 a = a + 1
 ReDim Preserve oValues(1 To a)
 oValues(a) = cell.Value
Next cell

CollectValues = oValues
End Function

Function sFindAllWireTypes()
Dim a As Integer, ws As Worksheet, b As Integer, oValues() As Variant
Set ws = ThisWorkbook.Worksheets("saved")

Do Until ws.Cells(a + 1, 1).Value = "Wire Name"
a = a + 1
If ws.Cells(a, 1).Value > "" Then
 b = b + 1
 ReDim Preserve oValues(1 To b)
 oValues(b) = ws.Cells(a, 1).Value
End If
Loop

sFindAllWireTypes = oValues

End Function
