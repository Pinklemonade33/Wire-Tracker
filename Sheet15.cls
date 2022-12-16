VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim exitdrop As Boolean



Private Sub Clear_Click()
Dim a As Integer, xrange As Integer, yrange As Integer
Dim r1 As Range, cell As Range, r2 As Range, r3 As Range

Set r1 = Range(Cells(7, 3), Cells(7, 30))
Set r2 = Range(Cells(8, 1), Cells(100, 30))
Set r3 = Range(Cells(2, 7), Cells(5, 30))

r1.Interior.ColorIndex = -4142
r1.Value = ""

r2.Value = ""

r3.Value = ""

End Sub

Private Sub cSettings_Click()
Settings.Show
End Sub

Private Sub cWireBox_Click()
Dim a As Integer
     
Do
a = a + 1
 If Cells(7, a).Value = OLEObjects("cWireBox").object.Value Then
  MsgBox ("Wire type already selected")
  
  Exit Sub
 End If
 If Cells(7, a).Value = "" Then Exit Do
Loop

With Cells(7, a)
 .Value = OLEObjects("cWireBox").object.Value
 .Interior.Color = RGB(0, 112, 192)
 .Font.Size = 14
 .Font.Color = -4142
End With

OLEObjects("cWireBox").object.Value = ""

End Sub



Private Sub CommandButton1_Click()
Dim cell As Range, InputList() As Variant, a As Integer, b As Integer, c As Integer, good As Boolean, prevlist() As Variant, ca As Boolean, Caa As Boolean

For Each cell In Selection
 If cell.Interior.Color = RGB(0, 112, 192) Then
  a = a + 1
  ReDim Preserve InputList(1 To a)
  InputList(a) = cell.Value
 End If
Next cell

Caa = checkarray(InputList)
If Caa = True Then
a = 6
Do Until Cells(2, a).Value = "" And a > 6
 a = a + 1
 c = c + 1
 ReDim Preserve prevlist(1 To c)
 prevlist(c) = Cells(2, a).Value
Loop

For b = 1 To UBound(InputList)
good = True
ca = checkarray(prevlist)
If ca = True Then
 For c = 1 To UBound(prevlist)
  If prevlist(c) = InputList(b) Then good = False
 Next c
End If
 If good = True Then
  Cells(2, a) = InputList(b)
  a = a + 1
 End If
If good = False Then MsgBox (InputList(b) & " is already selected")
Next b
End If
End Sub

Private Sub CommandButton2_Click()
Dim cell As Range, InputList() As Variant, a As Integer, b As Integer, c As Integer, good As Boolean, prevlist() As Variant, ca As Boolean, Caa As Boolean

For Each cell In Selection
  If cell.Interior.Color = RGB(0, 112, 192) Then
   a = a + 1
   ReDim Preserve InputList(1 To a)
   InputList(a) = cell.Value
 End If
Next cell

Caa = checkarray(InputList)
If Caa = True Then
a = 6
Do Until Cells(3, a).Value = "" And a > 6
 a = a + 1
 c = c + 1
 ReDim Preserve prevlist(1 To c)
 prevlist(c) = Cells(3, a).Value
Loop

For b = 1 To UBound(InputList)
good = True
ca = checkarray(prevlist)
If ca = True Then
 For c = 1 To UBound(prevlist)
  If prevlist(c) = InputList(b) Then good = False
 Next c
End If
 If good = True Then
  Cells(3, a) = InputList(b)
  a = a + 1
 End If
 If good = False Then MsgBox (InputList(b) & " is already selected")
Next b
End If

End Sub

Private Sub CommandButton3_Click()
Dim cell As Range, InputList() As Variant, a As Integer, b As Integer, c As Integer, good As Boolean, prevlist() As Variant, ca As Boolean, Caa As Boolean

For Each cell In Selection
 If cell.Interior.Color = RGB(0, 112, 192) Then
   a = a + 1
   ReDim Preserve InputList(1 To a)
   InputList(a) = cell.Value
 End If
Next cell

Caa = checkarray(InputList)
If Caa = True Then
a = 6
Do Until Cells(4, a).Value = "" And a > 6
 a = a + 1
 c = c + 1
 ReDim Preserve prevlist(1 To c)
 prevlist(c) = Cells(4, a).Value
Loop

For b = 1 To UBound(InputList)
good = True
ca = checkarray(prevlist)
If ca = True Then
 For c = 1 To UBound(prevlist)
  If prevlist(c) = InputList(b) Then good = False
 Next c
End If
 If good = True Then
  Cells(4, a) = InputList(b)
  a = a + 1
 End If
 If good = False Then MsgBox (InputList(b) & " is already selected")
Next b
End If

End Sub

Private Sub CommandButton4_Click()
Dim cell As Range, InputList() As Variant, a As Integer, b As Integer, c As Integer, good As Boolean, prevlist() As Variant, ca As Boolean, Caa As Boolean

For Each cell In Selection
 If cell.Interior.Color = RGB(0, 112, 192) Then
   a = a + 1
   ReDim Preserve InputList(1 To a)
   InputList(a) = cell.Value
 End If
Next cell

Caa = checkarray(InputList)
If Caa = True Then
a = 6
Do Until Cells(5, a).Value = "" And a > 6
 a = a + 1
 c = c + 1
 ReDim Preserve prevlist(1 To c)
 prevlist(c) = Cells(5, a).Value
Loop

For b = 1 To UBound(InputList)
good = True
ca = checkarray(prevlist)
If ca = True Then
 For c = 1 To UBound(prevlist)
  If prevlist(c) = InputList(b) Then good = False
 Next c
End If
 If good = True Then
  Cells(5, a) = InputList(b)
  a = a + 1
 End If
 If good = False Then MsgBox (InputList(b) & " is already selected")
Next b
End If

End Sub


Private Sub CreateGroup_Click()
    Call CreateGroup1
End Sub


Private Sub cWireBox_DropButtonClick()
Dim ws1 As Worksheet, a As Integer, b As Integer, ExistingWire As Boolean
If exitdrop = True Then Exit Sub
Set ws1 = ThisWorkbook.Worksheets("Saved")

ExistingWire = sCheckWire("Wire Name")
If ExistingWire = False Then MsgBox ("Selected wire does not exist")

Do Until cWireBox.ListCount = 0
 cWireBox.object.RemoveItem 0
Loop

Do
a = a + 1
 If ws1.Cells(a, 1).Value > "" Then
  b = b + 1
  If ws1.Cells(a, 1).Value = "Wire Name" Then Exit Do
  cWireBox.object.AddItem ws1.Cells(a, 1).Value
 End If
Loop

End Sub

Private Sub cWireBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim oValues() As Variant, iValues() As Variant, a As Integer, check As Boolean
If cWireBox.Value = "" Then Exit Sub
iValues = sFindAllWireTypes()
oValues = FindFullWord(cWireBox.Value, iValues)
cWireBox.Clear
check = checkarray(oValues)
If check = True Then
For a = 1 To UBound(oValues)
 cWireBox.AddItem oValues(a)
 Debug.Print oValues(a)
Next a
exitdrop = True
cWireBox.DropDown
End If
End Sub

Private Sub CommandButton34_Click()

Private Sub DeleteGroup_Click()
Dim a As Integer
    Call DeleteGroup1
    For a = 0 To OLEObjects("combobox3").object.ListCount + -1
        If OLEObjects("combobox3").object.list(a) = OLEObjects("combobox3").object.Value Then
           OLEObjects("combobox3").object.RemoveItem (a)
           OLEObjects("combobox3").object.Value = ""
           Exit For
        End If
    Next a
End Sub




Private Sub StartCalculation_Click()
Dim a As Integer, c As Integer, WireType() As Variant, setting() As String, check As Boolean

check = CheckSites
If check = False Then Exit Sub

a = 6
Do Until Cells(2, a + 1).Value = ""
 a = a + 1
 c = c + 1
 ReDim Preserve WireType(1 To c)
 ReDim Preserve setting(1 To c)
 WireType(c) = Cells(2, a).Value
 setting(c) = "StartOnStart"
Loop

a = 6
Do Until Cells(3, a + 1).Value = ""
 a = a + 1
 c = c + 1
 ReDim Preserve WireType(1 To c)
 ReDim Preserve setting(1 To c)
 WireType(c) = Cells(3, a).Value
 setting(c) = "StartOnCut"
Loop

a = 6
Do Until Cells(4, a + 1).Value = ""
 a = a + 1
 c = c + 1
 ReDim Preserve WireType(1 To c)
 ReDim Preserve setting(1 To c)
 WireType(c) = Cells(4, a).Value
 setting(c) = "StartOnTrim"
Loop

a = 6
Do Until Cells(5, a + 1).Value = ""
 a = a + 1
 c = c + 1
 ReDim Preserve WireType(1 To c)
 ReDim Preserve setting(1 To c)
 WireType(c) = Cells(5, a).Value
 setting(c) = "StartOnCalc"
Loop

Call Main(WireType, setting)

End Sub


Private Sub Select2_Click()
Dim a As Integer

For a = 0 To OLEObjects("combobox2").object.ListCount + -1
If OLEObjects("combobox2").object.Value = OLEObjects("combobox2").object.list(a) Then Call ComboBox2_Click
Next a

End Sub

Private Sub select3_Click()
Dim a As Integer

For a = 0 To OLEObjects("combobox3").object.ListCount + -1
If OLEObjects("combobox3").object.Value = OLEObjects("combobox3").object.list(a) Then Call SelectGroup
Next a

End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim a As Integer, r1 As Range, r2 As Range, good As Boolean, good2 As Boolean, b As Integer
Dim c As Integer, r2Array() As Variant, r1Array() As Variant, cell As Range, LastCol As Integer

good = True
a = 2
Do Until Cells(7, a + 1).Value = ""
  a = a + 1
Loop
 
If a = 2 Then good = False

If good = True Then
    LastCol = a
    Set r1 = Range(Cells(7, 3), Cells(7, a))
    If ActiveCell.Row = r1.Row And ActiveCell.Column <= LastCol Then
     For Each cell In r1
       If cell.Value <> ActiveCell.Value Then
        c = c + 1
        ReDim Preserve r1Array(1 To c)
        r1Array(c) = cell.Value
       End If
     Next cell
     Range(Cells(ActiveCell.Row, 3), Cells(ActiveCell.Row, LastCol)).Value = ""
     Range(Cells(ActiveCell.Row, 3), Cells(ActiveCell.Row, LastCol)).Interior.ColorIndex = -4142
     If c > 0 Then
      c = 1
      Do Until c = UBound(r1Array) + 1
       With Cells(ActiveCell.Row, 2 + c)
        .Value = r1Array(c)
        .Interior.Color = RGB(0, 112, 192)
        c = c + 1
       End With
      Loop
     End If
    End If
End If

good = True
good2 = True
a = 2
b = 6
Do Until good2 = False
good2 = False
a = 2
b = b + 1
 Do Until a = 6
  If Cells(a, b).Value > "" Then good2 = True
  a = a + 1
 Loop
Loop
If b = 7 And good2 = False Then good = False

If good = True Then
    Set r2 = Range(Cells(2, 7), Cells(a, b))
    LastCol = b
    If ActiveCell.Column <= LastCol And ActiveCell.Column >= 7 Then
     If ActiveCell.Row >= 2 And ActiveCell.Row <= 5 Then
      For Each cell In r2
       If cell.Row = ActiveCell.Row And cell.Value <> ActiveCell.Value Then
        c = c + 1
        ReDim Preserve r2Array(1 To c)
        r2Array(c) = cell.Value
       End If
     Next cell
     Range(Cells(ActiveCell.Row, 7), Cells(ActiveCell.Row, LastCol)).Value = ""
     If c > 0 Then
      c = 1
      Do Until c = UBound(r2Array) + 1
       Cells(ActiveCell.Row, 6 + c).Value = r2Array(c)
       c = c + 1
      Loop
     End If
    End If
  End If
End If

End Sub

Function CheckSites()
Dim a As Integer, b As Integer, check As Boolean, iValues() As Variant, oValues() As Variant, answer As String, col(1 To 2), c As Integer
Dim cStrings(1 To 2) As String, Joined As String

col(1) = 1
col(2) = 2
cStrings(1) = "Job#'s:"
cStrings(2) = "Site ID's:"


Do Until c = 2
c = c + 1
Erase iValues
Erase oValues
b = 0
a = 7
 Do Until Cells(a + 1, col(c)).Value = ""
  a = a + 1
  b = b + 1
  ReDim Preserve iValues(1 To b)
  iValues(b) = Cells(a, col(c)).Value
 Loop
oValues = FindDuplicates(iValues)
check = checkarray(oValues)
If check = True Then
 Joined = Join(oValues, " ")
 answer = MsgBox(cStrings(c) & " " & Joined & " Are listed twice would you like to continue anyways?", vbYesNo)
 If answer = vbNo Then
  CheckSites = False
  Exit Function
 End If
End If
Loop
 
CheckSites = True

End Function

Sub test()

With OLEObjects("ManualList").object
 .AddItem
 .list(0, 1) = 1
 .list(0, 2) = 2
End With

End Sub

