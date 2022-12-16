Attribute VB_Name = "Module6"
Option Explicit
Dim ActiveRow As Integer, ActiveCol As Integer, LowestDiff As Integer, SpecMult() As Variant

Sub CalculateTransfer(SpecList, BaseList, Amount, Trim, Over, Under)
Dim a As Integer, b As Integer, cell As Range, c As Integer
Dim FirstCutsRange As Integer, LastCutsRange As Integer
Dim FirstSpoolsRange As Integer, LastSpoolsRange As Integer
Dim FirstPreCutsRange As Integer, LastPreCutsRange As Integer
Dim FirstInactiveSpoolsRange As Integer, LastInactiveSpoolsRange As Integer
Dim FirstSelectedPreCutsRange As Integer, LastSelectedPreCutsRange As Integer
Dim ActiveSpoolRange As Integer, LastRow As Integer, CellRange As Range, NumberOfCuts As Integer
Dim CR As Range, SR As Range, PR As Range, ASR As Range, SPR As Range, ISR As Range, BIR As Range
Dim SpoolsAndCuts() As Integer, AllCombinations() As Variant, PrevNoc As Integer, SpecArr() As Integer
Dim LowestBase As Integer, BaseIncrements() As Integer, AmountNeeded As Integer, sel As Boolean, SpecMult() As Variant
Dim CurrentFirstRange As Integer, cRanges() As Integer, CurrentSum As Integer, PrevSum As Integer, cSums() As Integer
Dim tp As Worksheet, MN As Worksheet

Set tp = ThisWorkbook.Worksheets("Tempsave")
Set MN = ThisWorkbook.Worksheets("Manual")

Do Until MN.Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
Loop
LastRow = a

a = 0
Do Until MN.Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
 With MN.Cells(a, 1)
  If .Value = "Requested Cuts" Then FirstCutsRange = a
  If .Value = "Spools" Then
   FirstSpoolsRange = a + 1
   LastCutsRange = a + -1
  End If
  If .Value = "Pre-Cuts" Then
   FirstPreCutsRange = a + 1
   LastSpoolsRange = a + -1
  End If
 End With
 If MN.Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then LastPreCutsRange = a + 1
Loop

a = 0
Do Until MN.Cells(1, a + 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
 If MN.Cells(1, a).Value = "Selected Pre-Cuts" Then FirstSelectedPreCutsRange = a
 If MN.Cells(1, a).Value = "Actions" Then
    LastSelectedPreCutsRange = a + -2
    Set BIR = Range(MN.Cells(7, a), MN.Cells(LastRow, a + 2))
 End If
Loop

a = 0
LastInactiveSpoolsRange = FirstSelectedPreCutsRange + -2


ActiveSpoolRange = 8
FirstInactiveSpoolsRange = ActiveSpoolRange + 2
LastInactiveSpoolsRange = FirstSelectedPreCutsRange + -2

Set CR = Range(MN.Cells(FirstCutsRange, 1), MN.Cells(LastCutsRange, 6))
Set SR = Range(MN.Cells(FirstSpoolsRange, 1), MN.Cells(LastSpoolsRange, 6))
Set PR = Range(MN.Cells(FirstPreCutsRange, 1), MN.Cells(LastPreCutsRange, 6))
Set ASR = Range(MN.Cells(3, 8), MN.Cells(LastRow, 8))
Set SPR = Range(MN.Cells(2, FirstSelectedPreCutsRange), MN.Cells(LastRow, LastSelectedPreCutsRange))
Set ISR = Range(MN.Cells(3, FirstInactiveSpoolsRange), MN.Cells(3, LastInactiveSpoolsRange))

For Each cell In BIR
 If cell.Value > "" Then
  c = c + 1
  ReDim Preserve BaseIncrements(1 To c)
  BaseIncrements(c) = cell.Value
  If cell.Value < LowestBase Or c = 1 Then LowestBase = cell.Value
 End If
Next cell

c = 0
For Each cell In SR
 If cell.Value > "" And cell.Value >= LowestBase Then
  c = c + 1
  ReDim Preserve SpoolsAndCuts(1 To c)
  SpoolsAndCuts(c) = cell.Value
 End If
Next cell
c = 0
For Each cell In PR
 If cell.Value > "" And cell.Value >= LowestBase Then
  c = c + 1
  ReDim Preserve SpoolsAndCuts(1 To c)
  SpoolsAndCuts(c) = cell.Value
 End If
Next cell

AllCombinations = FindAllCombinations(SpoolsAndCuts, AmountNeeded, AmountNeeded)
NumberOfCuts = FindNumberOfCuts(AllCombinations)

With tCombinationList
 .tQTYbox.object.AddItem NumberOfCuts
 .Show
End With

Do
DoEvents

If tCombinationList.tQTYbox.Value <> PrevNoc Then
 Do Until tCombinationList.tSumBox.object.ListCount = 0
  tCombinationList.tSumBox.object.RemoveItem 0
 Loop
 c = 0
 Erase cRanges
 Erase cSums
 For a = 1 To UBound(AllCombinations)
  If AllCombinations(a) = "Cuts" Then CurrentFirstRange = a
  If AllCombinations(a) = "Sum" Then CurrentSum = AllCombinations(a + 1)
  If AllCombinations(a) = "Number Of Cuts" Then
   If AllCombinations(a + 1) = tCombinationList.tQTYbox.Value Then
    c = c + 1
    ReDim Preserve cRanges(1 To c)
    ReDim Preserve cSums(1 To c)
    cRanges(c) = CurrentFirstRange
    cSums(c) = CurrentSum
    tCombinationList.tSumBox.object.AddItem CurrentSum
   End If
  End If
 Next a
End If

PrevNoc = tCombinationList.tQTYbox.Value
If tCombinationList.tSumBox.Value <> PrevSum Then
 With tCombinationList.tCutBox
  Do Until .object.ListCount = 0
   .object.RemoveItem 0
  Loop
  a = cRanges(.object.ListIndex + 1)
 End With
 
  Do Until AllCombinations(a) = "Sum"
   a = a + 1
   sel = False
   For Each cell In PR
    If cell.Value = AllCombinations(a) Then
     Set CellRange = Cells(cell.Row, cell.Column)
     Call FindActiveCell(cell.Value, CellRange)
     Call Pick(0, 255, 0, FirstCutsRange, LastCutsRange, ActiveRow, ActiveCol, 2, LastRow, FirstSelectedPreCutsRange, LastSelectedPreCutsRange, False)
     sel = True
     Exit For
    End If
   Next cell
   If sel = False Then
   For Each cell In SR
    If cell.Value = AllCombinations(a) Then
     Set CellRange = Cells(cell.Row, cell.Column)
     Call FindActiveCell(cell.Value, CellRange)
     Call Pick(0, 255, 0, FirstCutsRange, LastCutsRange, ActiveRow, ActiveCol, 2, LastRow, FirstSelectedPreCutsRange, LastSelectedPreCutsRange, False)
     Exit For
    End If
   Next cell
   End If
  Loop
 End If
Loop



End Sub

Sub FindActiveCell(CellValue, CellRange)
Dim cell As Range

For Each cell In CellRange
 If cell.Value = CellValue And cell.Interior.Color = RGB(255, 255, 255) Then
  ActiveRow = cell.Row
  ActiveCol = cell.Column
 End If
Next cell

End Sub

Function FindNumberOfCuts(combo)
Dim a As Integer, b As Integer, NumberOfCuts() As Integer, First As Boolean

First = True
For a = 1 To UBound(combo)
 If combo(a) = "Number Of Cuts" Then
  good = True
  If First = False Then
   For b = 1 To UBound(NumberOfCuts)
    If NumberOfCuts(b) = combo(a + 1) Then good = False
   Next b
  End If
  If good = True Then
   c = c + 1
   ReDim Preserve NumberOfCuts(1 To c)
   NumberOfCuts(c) = combo(a + 1)
   First = False
  End If
 End If
Next a

FindNumberOfCuts = NumberOfCuts

End Function

Sub PopulateSpecMult(InputSpec, InputDiff)
Dim a As Integer, b As Integer, cell As Range, DivList() As Integer
Dim HighestNumber As Integer, LowestNumber As Integer, Qty As Integer

ReDim DivList(1 To UBound(InputSpec))

For Each cell In InputSpec
 If cell.Value > HighestNumber Then HighestNumber = cell.Value
Next cell

LowestNumber = HighestNumber
For Each cell In InputSpec
 If cell.Value < LowestNumber Then LowestNumber = cell.Value
Next cell

For a = 1 To InputSpec.Rows.Count
 If InputDiff < InputSpec.Rows(a).Value Then
  DivList(a) = InputDiff \ InputSpec.Rows(a).Value
 End If
 Qty = Qty + DivList(a)
Next a

ReDim SpecMult(1 To Qty)

For a = 1 To UBound(DivList)
 For b = 1 To DivList(a)
  c = c + 1
  SpecMult(c) = InputSpec.Rows(a).Value
 Next b
Next a

SpecMult = FindAllCombinations(SpecMult, InputDiff, InputDiff)
SpecMult = FindOptimalSpecCombination(SpecMult, InputDiff, Trim.Value, Thresh.Value, MaxThresh.Value, InputSpec, InputQTY, False)
   
End Sub

