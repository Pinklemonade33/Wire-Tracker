Attribute VB_Name = "Module7"
Option Explicit

Sub ClearManual()
Dim a As Integer, MN As Worksheet
Dim LastCol As Integer, LastRow As Integer, good As Boolean, b As Integer, good2 As Boolean

Set MN = ThisWorkbook.Worksheets("Manual")

a = 1
Do Until MN.Cells(1, a).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
Loop
    LastCol = a
   
a = 0
Do Until good = True
good2 = False
  Do Until good2 = True
   a = a + 1
   If MN.Cells(a, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then good2 = True
  Loop
  good = True
  b = 0
  Do Until b = LastCol
   b = b + 1
   If Not MN.Cells(a, b).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then good = False
  Loop
Loop

    LastRow = a
    
Range(MN.Cells(1, 1), MN.Cells(LastRow, LastCol)).Delete
Range(MN.Columns(1), MN.Columns(LastCol)).ColumnWidth = 8
MN.Buttons.Delete

End Sub

Function FindCutRanges(InputResults, InputLap, InputKeyWord)
Dim a As Integer, HighestNumber As Integer


For a = LBound(InputResults) To UBound(InputResults)
  If InputResults(a) = "Lap" Then
    If InputResults(a + 1) = InputLap Then
   Do Until InputResults(a) = InputKeyWord
    a = a + 1
   Loop
   If InputResults(a + 1) > HighestNumber Then HighestNumber = InputResults(a + 1)
    End If
  End If
   
Next a

FindCutRanges = HighestNumber

End Function

Sub ProjectColors(FirstRange, LastRange, Red, Green, Blue, Size, InputBold, InputValue, InputHeight, Header, inputlastcol)
Dim pj As Worksheet

Set pj = ThisWorkbook.Worksheets("Projection")

Range(pj.Cells(FirstRange, 1), pj.Cells(LastRange, inputlastcol)).Interior.Color = RGB(Red, Green, Blue)
With pj.Cells(FirstRange, 1)
 .Font.Color = 1
 .Font.Size = Size
 .Font.Bold = InputBold
 .Value = InputValue
 .RowHeight = InputHeight
 If Header = False Then
  With .Borders(xlEdgeRight)
  .LineStyle = Excel.XlLineStyle.xlContinuous
  .Weight = xlThick
  End With
 End If
  If FirstRange < LastRange Then .Borders(xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
End With

If Header = True Then
With Range(pj.Cells(FirstRange, 1), pj.Cells(FirstRange, inputlastcol))
  .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
   With .Borders(xlEdgeTop)
    .LineStyle = Excel.XlLineStyle.xlContinuous
    .Weight = xlThick
   End With
   With .Borders(xlEdgeBottom)
    .LineStyle = Excel.XlLineStyle.xlContinuous
    .Weight = xlThick
   End With
End With
 pj.Cells(FirstRange, inputlastcol).Borders(xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
End If
 
If FirstRange < LastRange Then
With Range(pj.Cells(FirstRange + 1, 1), pj.Cells(LastRange, 1))
  .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
  With .Borders(xlEdgeRight)
   .LineStyle = Excel.XlLineStyle.xlContinuous
   .Weight = xlThick
  End With
End With
pj.Cells(LastRange, 1).Borders(xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
End If

End Sub


Function FindHighestCut(InputCuts, InputCutStatus, InputKeyWord)
Dim a As Integer, CurrentCut As Integer, HighestCut As Integer

For a = LBound(InputCuts) To UBound(InputCuts)
  If InputCuts(a) = "Type" Then
   a = a + 1
    Do Until InputCuts(a + 1) = "end"
     a = a + 1
     If InputCutStatus(a) = InputKeyWord Then CurrentCut = CurrentCut + 1
    Loop
      If HighestCut < CurrentCut Then HighestCut = CurrentCut
      CurrentCut = 0
  End If
Next a


FindHighestCut = HighestCut

End Function

Function FindHighestAdd(InputAddCuts)
Dim a As Integer, CurrentCut As Integer, HighestCut As Integer

For a = LBound(InputAddCuts) To UBound(InputAddCuts)
  If InputAddCuts(a) = "Type" Then
   a = a + 1
    Do Until InputAddCuts(a + 1) = "end"
     a = a + 1
     CurrentCut = CurrentCut + 1
    Loop
      If HighestCut < CurrentCut Then HighestCut = CurrentCut
      CurrentCut = 0
  End If
Next a

FindHighestAdd = HighestCut

End Function


Function FindAdditionalCuts(InputTypeList, InputResults)
Dim a As Integer, b As Integer, c As Integer, AddCuts() As Variant
Dim First As Boolean, FirstRange As Integer


For a = LBound(InputTypeList) To UBound(InputTypeList)
First = True
FirstRange = 0
  For b = LBound(InputResults) To UBound(InputResults)
   If First = True Then
    c = c + 1
    ReDim Preserve AddCuts(1 To c + 1)
    AddCuts(c) = "Type"
    c = c + 1
    AddCuts(c) = InputTypeList(a)
    First = False
    FirstRange = c
   End If
    If InputResults(b) = InputTypeList(a) Then
       Do Until InputResults(b) = "Pre-Cuts"
         b = b + 1
       Loop
       If InputResults(b + 1) > 0 Then
        Do Until InputResults(b + 1) = "end"
          b = b + 1
          c = c + 1
          ReDim Preserve AddCuts(1 To c)
          AddCuts(c) = InputResults(b)
        Loop
       End If
    End If
  Next b
 
  c = c + 1
  ReDim Preserve AddCuts(1 To c)
  AddCuts(c) = "end"
 
 If FirstRange = c Then
  c = c + 1
  ReDim Preserve AddCuts(1 To c)
  AddCuts(c) = 0
 End If

Next a

FindAdditionalCuts = AddCuts

End Function

Function GetSpools(InputWireType, InputSpoolResults)
Dim a As Integer, b As Integer, Spools() As Integer

For a = 1 To UBound(InputSpoolResults)
 If InputSpoolResults(a) = "Type" Then
   a = a + 1
    If InputSpoolResults(a) = InputWireType Then
      a = a + 4
      Do Until InputSpoolResults(a + 1) = "end"
       a = a + 1
       b = b + 1
       ReDim Preserve Spools(1 To b)
       If InputSpoolResults(a) <> 0 Then Spools(b) = InputSpoolResults(a)
      Loop
    End If
 End If
Next a

GetSpools = Spools
End Function

Function GetCuts(AllCuts, InputCutStatus, InputWireType)
Dim a As Integer, b As Integer, Cuts() As Integer

For a = 1 To UBound(AllCuts)
 If AllCuts(a) = "Type" Then
   a = a + 1
    If AllCuts(a) = InputWireType Then
      Do Until AllCuts(a + 1) = "end"
       a = a + 1
       b = b + 1
       ReDim Preserve Cuts(1 To b)
       Cuts(b) = AllCuts(a)
      Loop
    End If
 End If
Next a

GetCuts = Cuts
End Function

Function GetPreCuts(InputHighCuts, InputLowCuts, InputWireType)
Dim a As Integer, b As Integer, PreCuts() As Integer

For a = 1 To UBound(InputHighCuts)
 If InputHighCuts(a) = "Type" Then
   a = a + 1
   If InputHighCuts(a) = InputWireType Then
    Do Until InputHighCuts(a + 1) = "end"
     a = a + 1
     b = b + 1
     ReDim Preserve PreCuts(1 To b)
     PreCuts(b) = InputHighCuts(a)
    Loop
   End If
 End If
Next a

For a = 1 To UBound(InputLowCuts)
 If InputLowCuts(a) = "Type" Then
   a = a + 1
   If InputLowCuts(a) = InputWireType Then
    Do Until InputLowCuts(a + 1) = "end"
     a = a + 1
     b = b + 1
     ReDim Preserve PreCuts(1 To b)
     PreCuts(b) = InputLowCuts(a)
    Loop
   End If
 End If
Next a

GetPreCuts = PreCuts
End Function

Sub ManualColors(FirstRange, LastRange, Red, Green, Blue, Size, InputBold, InputValue, InputHeight, Header, inputlastcol, inputfirstcol, WhiteFont)
Dim MN As Worksheet

Set MN = ThisWorkbook.Worksheets("Manual")

Range(MN.Cells(FirstRange, inputfirstcol), MN.Cells(LastRange, inputlastcol)).Interior.Color = RGB(Red, Green, Blue)
If Not InputValue = "" Then MN.Cells(FirstRange, inputfirstcol).Value = InputValue
With Range(MN.Cells(FirstRange, inputfirstcol), MN.Cells(LastRange, inputlastcol))
 If WhiteFont = True Then .Font.Color = -4142
 If WhiteFont = False Then .Font.Color = 1
 .Font.Size = Size
 .Font.Bold = InputBold
 If InputValue = "" Then .Value = InputValue
 .RowHeight = InputHeight
End With

If Header = True Then
MN.Cells(FirstRange, inputfirstcol).Value = InputValue
With Range(MN.Cells(FirstRange, inputfirstcol), MN.Cells(FirstRange, inputlastcol))
  .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
   With .Borders(xlEdgeTop)
    .LineStyle = Excel.XlLineStyle.xlContinuous
    .Weight = xlThick
   End With
   With .Borders(xlEdgeBottom)
    .LineStyle = Excel.XlLineStyle.xlContinuous
    .Weight = xlThick
   End With
End With
End If
 
If Header = False Then
With Range(MN.Cells(FirstRange, inputfirstcol), MN.Cells(LastRange, inputlastcol))
  .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
End With
End If

End Sub

Sub ManualBorder(InputCol, InputLastRow, InputWidth, InputFirstRow)
Dim a As Integer
Dim MN As Worksheet
Set MN = ThisWorkbook.Worksheets("Manual")

MN.Columns(InputCol).ColumnWidth = InputWidth
With Range(MN.Cells(InputFirstRow, InputCol), MN.Cells(InputLastRow, InputCol))
  .Interior.Color = RGB(119, 119, 119)
  .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
  .Borders(xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
  .Borders(xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
  .Borders(xlEdgeRight).Weight = xlThick
  .Borders(xlEdgeLeft).Weight = xlThick
End With

End Sub

Sub ManualBorder2(inputfirstcol, inputrow, InputHeight, inputlastcol)
Dim a As Integer, MN As Worksheet
Set MN = ThisWorkbook.Worksheets("Manual")

MN.Rows(inputrow).RowHeight = InputHeight
With Range(MN.Cells(inputrow, inputfirstcol), MN.Cells(inputrow, inputlastcol))
 .Interior.Color = RGB(119, 119, 119)
 .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 .Borders(xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
 .Borders(xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
 .Borders(xlEdgeTop).Weight = xlThick
 .Borders(xlEdgeBottom).Weight = xlThick
End With

End Sub


Sub HighlightWire(AllCuts, CutStatus, SpoolCutRange, WireType, FirstCutRange, LastCutRange, SelectedSpoolRange, FirstSelectedPreCutRange, SpoolResults, LastSelectedPreCutRange, FirstPreCutRange)
Dim a As Integer, b As Integer, c As Integer, MN As Worksheet, SpoolToChart() As Integer, SpoolRangeList() As Integer
Dim red1 As Integer, green1 As Integer, blue1 As Integer, good As Boolean, e As Integer, d As Integer
Set MN = ThisWorkbook.Worksheets("Manual")
ReDim SpoolToChart(1 To UBound(AllCuts))
ReDim SpoolRangeList(1 To 1)

red1 = 0
green1 = 176
blue1 = 80

e = 2
c = 1
For a = 1 To UBound(AllCuts)
  If AllCuts(a) = "Type" Then
   a = a + 1
   If AllCuts(a) = WireType Then
    Do Until AllCuts(a) = "end"
     a = a + 1
     If CutStatus(a) = "Pre-Cut" Then
      b = FirstCutRange
      Do Until MN.Cells(b, c).Value = AllCuts(a) And Not MN.Cells(b, c).Interior.ColorIndex = 4
       If c = 12 Then
        c = 0
        b = b + 1
       End If
       c = c + 1
      Loop
      d = FirstSelectedPreCutRange
      e = 2
      Do Until MN.Cells(e, d).Value = ""
       d = d + 1
       If d = LastSelectedPreCutRange Then
        d = FirstSelectedPreCutRange
        e = e + 1
       End If
      Loop
       MN.Cells(e, d).Value = AllCuts(a)
       MN.Cells(e, d).Interior.ColorIndex = 4
       MN.Cells(b, c).Interior.ColorIndex = 4
       b = FirstPreCutRange
      Do Until MN.Cells(b, c).Value = AllCuts(a) And Not MN.Cells(b, c).Interior.Color = 4
       If c = 12 Then
        c = 0
        b = b + 1
       End If
       c = c + 1
      Loop
      MN.Cells(b, c).Interior.ColorIndex = 4
     End If
    Loop
  End If
 End If
Next a
          
For a = 1 To UBound(AllCuts)
 If AllCuts(a) = "Type" Then
  a = a + 1
   If AllCuts(a) = WireType Then
    Do Until AllCuts(a) = "end"
    a = a + 1
    If SpoolCutRange(a) > "" Then
     good = False
     For b = 1 To UBound(SpoolRangeList)
      If SpoolRangeList(b) <> SpoolCutRange(a) Then good = True
     Next b
     If good = True Then
     c = c + 1
     ReDim Preserve SpoolRangeList(1 To c)
     SpoolRangeList(c) = SpoolCutRange(a)
     End If
    End If
   Loop
  End If
 End If
Next a

e = SelectedSpoolRange + -1
For a = 1 To UBound(SpoolRangeList)
d = 5
 For b = 1 To UBound(SpoolResults)
  If b = SpoolRangeList(a) Then
   e = e + 1
   MN.Cells(3, e).Value = SpoolResults(b)
   For c = 1 To UBound(AllCuts)
    If SpoolCutRange(c) = SpoolRangeList(a) Then
     MN.Cells(d, e).Value = AllCuts(c)
     MN.Cells(d, e).Interior.Color = RGB(red1, green1, blue1)
     d = d + 1
    End If
   Next c
  End If
 Next b
Next a

End Sub

Sub CheckSelected(FirstSelectedRange, LastSelectedRange, FirstBorderRange, SecondBorderRange, ThirdBorderRange, FirstSelectedSpoolRange, _
LastSelectedSpoolRange, FirstSelectedPreCutRange, LastSelectedPreCutRange, red2, green2, blue2)

Dim a As Integer, b As Integer, good As Boolean, MN As Worksheet

Set MN = ThisWorkbook.Worksheets("Manual")

a = 7
good = False
Do Until good = True
b = FirstSelectedRange
a = a + 1
good = True
 Do Until b = LastSelectedRange
  b = b + 1
  If MN.Cells(a, b).Value > "" And MN.Cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
   Call ManualBorder(FirstBorderRange, a, 3, 1)
   Call ManualColors(4, a, red2, green2, blue2, 14, False, "", 30, False, LastSelectedSpoolRange, FirstSelectedSpoolRange, False)
   Call ManualBorder(SecondBorderRange, a, 1, 2)
   Call ManualColors(3, a, red2, green2, blue2, 14, False, "", 30, False, LastSelectedPreCutRange, FirstSelectedPreCutRange, False)
   Call ManualBorder(ThirdBorderRange, a, 3, 1)
   good = False
  End If
 Loop
Loop

End Sub

Sub Borders1(InputFirstRow, InputLastRow, InputCol)
Dim MN As Worksheet
Set MN = ThisWorkbook.Worksheets("Manual")

With Range(MN.Cells(InputFirstRow, InputCol), MN.Cells(InputLastRow, InputCol))
 .Borders(xlEdgeRight).Weight = xlThick
End With

End Sub


Sub ExtendSpools(Extend)
Dim a As Integer, MN As Worksheet, cell As Range
Dim Border1 As Integer, Border2 As Integer, Border3 As Integer
Dim FirstInactiveRange As Integer, LastSpoolRange As Integer
Dim FirstPreCutsRange As Integer, LastPreCutsRange As Integer
Dim FirstActionsRange As Integer, LastActionsRange As Integer
Dim LastPreCutRow As Integer, LastRow As Integer, FirstSpoolRange As Integer
Set MN = ThisWorkbook.Worksheets("Manual")

FirstInactiveRange = 10

a = 7
Do Until MN.Cells(1, a).Value = ""
 a = a + 1
 If MN.Cells(1, a).Value = "Selected Pre-Cuts" Then
   FirstPreCutsRange = a
   Border1 = a + -1
 End If
 If MN.Cells(1, a).Value = "Actions" Then
  FirstActionsRange = a
  Border2 = a + -1
  LastPreCutsRange = a + -2
  Border3 = a + 4
  LastActionsRange = a + 7
 End If
Loop

a = 0
Do Until MN.Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
Loop
 LastRow = a
 
a = 0
Do Until MN.Cells(a + 1, FirstPreCutsRange).Value = ""
a = a + 1
Loop
 LastPreCutRow = a

ReDim PreCuts(1 To LastPreCutsRow * ((LastPreCutRange + -FirstPreCutsRange) + 1))

For Each cell In Range(MN.Cells(2, FirstPreCutsRange), MN.Cells(LastPreCutRow, LastPreCutsRange))
 PreCuts(a) = cell.Value
Next cell

Range(MN.Cells(1, Border1), MN.Cells(LastRow, LastActive)).Delete
MN.butons.Delete

If Extend = True Then b = 1
If Extend = False Then b = -1

Border1 = Border1 + b
Border2 = Border2 + b
Border3 = Border3 + b
LastSpoolRange = LastSpoolRange + b
FirstSpoolRange = 8
FirstPreCutsRange = FirstPreCutsRange + b
LastPreCutsRange = LastPreCutsRange + b
FirstActionsRange = FirstActionsRange + b
LastActionsRange = LastActionsRange + b

Call ManualColors(1, 1, 0, 32, 96, 18, True, "Selected Spools", 30, True, LastSpoolRange, FirstSpoolRange, True)
Call ManualColors(2, 2, 0, 176, 240, 14, True, "Inactive", 30, True, LastSpoolRange, FirstInactiveRange, True)
Call ManualColors(3, 3, 0, 32, 96, 14, False, "", 30, True, LastSpoolRange, FirstInactiveRange, False)
Call ManualColors(4, 4, 155, 194, 230, 14, False, "", 30, True, LastSpoolRange, FirstInactiveRange, False)
Call ManualColors(5, LastRow, 255, 255, 255, 14, False, "", 30, False, LastSpoolRange, FirstInactiveRange, False)
Call ManualBorder(Border1, LastRow, 1, 1)

Call ManualColors(1, 1, 0, 32, 96, 18, True, "Selected Pre-Cuts", 30, True, LastPreCutsRange, FirstPreCutsRange, True)
Call ManualColors(2, LastRow, 255, 255, 255, 14, False, "", 30, False, LastPreCutsRange, FirstPreCutsRange, False)
Call ManualBorder(Border2, LastRow, 3, 1)

Call ManualColors(1, 1, 0, 32, 96, 28, True, "Actions", 30, True, LastActionsRange, FirstActionsRange, True)
Call ManualColors(3, 3, 0, 32, 96, 14, True, "Transfer Calculator", 30, True, FirstActionsRange + 2, FirstActionsRange, True)
Call ManualColors(5, 5, 255, 255, 255, 14, False, "", 30, False, FirstActionsRange + 2, FirstActionsRange + 2, False)
Call ManualColors(7, LastRow, 255, 255, 255, 14, False, "", 30, False, FirstActionsRange + 2, FirstActionsRange, False)
Call ManualBorder(Border3, LastRow, 1, 3)
Call AddButton(4, Border3 + 1, Border3 + 3, "IncrementButton", "Start", "Manual")
Call AddButton(2, firstactionrange, lastactionrange + 1, "DoneButton", "Done", "Manual")

Call ManualColors(3, 3, 0, 32, 96, 14, True, "Increment Calculator", 30, True, Border3 + 1, Border3 + 3, True)
Call ManualColors(5, 5, 255, 255, 255, 14, False, "", 30, False, Border3 + 3, Border3 + 3, False)
Call ManualColors(6, 6, 255, 255, 255, 14, False, "", 30, False, Border3 + 3, Border3 + 3, False)
Call AddButton(4, FirstActionsRange, FirstActionsRange + 2, "TransferButton", "Start""Manual")

End Sub

