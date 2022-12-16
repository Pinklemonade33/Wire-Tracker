Attribute VB_Name = "mCalculate"
Option Explicit
Dim aa As Integer, Finished As Boolean, HighCuts() As Variant, SpoolCutRange() As Integer
Dim CurrentSpoolRange As Integer, LowCuts() As Variant, ActiveSpoolRange() As Integer
Dim NumberOfCuts As Integer, SpoolArray() As Variant, ActiveCutsRange As Range, NoSpec As Boolean
Dim SpoolQTY() As Integer, CutStatus() As String, SpoolStatus() As String, ProjectionRanges() As Integer
Dim PrefMaxTrim As Integer, CalculationList() As Variant, ActiveSpoolArray() As Variant
Dim CalculationStart As Integer, AllCutCombinations() As Variant, SpoolResults As Variant
Dim grouping As Integer, OptimalSpecCuts() As Integer, AllSpecCombinations() As Integer
Dim TotalSum As Integer, ActiveSpool() As Integer, Halt As Boolean, InputDiffLoc() As Integer, InputDiffQTY() As Integer, Cutslist() As Integer, LowSpec As Integer
Dim SpecMult() As Variant, LowQTY As Boolean, CutResults() As Variant, TypeList() As Variant
Dim OptimalSpecCombination() As Variant, DeSpool() As Variant
Dim InputDiff() As Integer, InputSpec() As Integer, AllCuts() As Variant, Lap As Integer
Dim StartOnStart() As Boolean, StartOnCut() As Boolean, StartOnTrim() As Boolean, StartOnCalc() As Boolean, ConLowCuts() As Variant, ConHighCuts() As Variant, OnlyPreCuts As Boolean
Dim SpecificCuts() As Integer, BaseIncrements() As Integer, qThresholds() As Integer, qMaxThresholds() As Integer, ThresholdStages() As Variant, Mode As String, ManualLaunch() As Variant
Dim mStandard As Boolean, mCritical As Boolean, mFree As Boolean, bb As Integer
Sub Main(WireType, setting)
SetCalculationList
Call ManualType(WireType, setting)
CountBulk
AdjustBulk
SetSpoolArray
SetActiveSpool
CollectCuts
CollectPreCuts
FindPreCuts

Do Until Finished = True
check
align
If Finished = True Then Exit Do
identify
If Halt = True Then Exit Sub
If StartOnStart(aa) = True Or StartOnCut(aa) = True Then Manual
If Not StartOnStart(aa) = True Then
DefineCuts
NewSpool
FindDifference
If NoSpec = True Then AltSpec
If NoSpec = False Then FindAllMultiples
Store
End If
Loop

CheckPreCuts
ClearProjection
If OnlyPreCuts = True Then ProjectOnlyPreCuts
If OnlyPreCuts = False Then
FindProjectionRanges
ProjectResults1
ProjectResults2
End If

End Sub

Sub ManualType(WireType, setting)
Dim a As Integer, b As Integer, checkb As Boolean

ReDim StartOnCut(1 To UBound(CalculationList))
ReDim StartOnStart(1 To UBound(CalculationList))
ReDim StartOnTrim(1 To UBound(CalculationList))
ReDim StartOnCalc(1 To UBound(CalculationList))

checkb = checkarray(setting)
If checkb = True Then
For a = 1 To UBound(WireType)
 For b = 1 To UBound(TypeList)
  If TypeList(b) = WireType(a) Then
    If setting(a) = "StartOnCut" Then StartOnCut(b) = True
    If setting(a) = "StartOnStart" Then StartOnStart(b) = True
    If setting(a) = "StartOnTrim" Then StartOnTrim(b) = True
    If setting(a) = "StartOnCalc" Then StartOnCalc(b) = True
  End If
 Next b
Next a
End If


End Sub

Sub check()
Dim a As Integer, b As Integer, c As Integer, good As Boolean, CurrentType As Integer, CurrentB As Integer

For a = LBound(SpoolArray) To UBound(SpoolArray)
  If SpoolArray(a) = "Type" Then
    For b = LBound(AllCuts) To UBound(AllCuts)
      If AllCuts(b) = "Type" Then
        If AllCuts(b + 1) = SpoolArray(a + 1) Then
        a = a + 4
        CurrentB = b + 2
         Do Until SpoolArray(a + 1) = "end"
           a = a + 1
           good = False
           b = CurrentB
           Do Until AllCuts(b) = "end"
             If SpoolResults(a) > AllCuts(b) And CutStatus(b) = "" Then good = True
             b = b + 1
           Loop
             If good = False Then SpoolStatus(a) = "Done"
         Loop
        End If
      End If
    Next b
  End If
Next a


For a = LBound(CalculationList) To UBound(CalculationList)
  If Not CalculationList(a) = "Done" Then
    For b = LBound(SpoolArray) To UBound(SpoolArray)
       If SpoolArray(b) = "Type" Then
        CurrentType = b + 1
        b = b + 4
        good = False
          Do Until SpoolArray(b + 1) = "end"
             b = b + 1
            If Not SpoolStatus(b) = "Done" Then good = True
          Loop
           If good = False Then CalculationList(a) = "Done"
        End If
    Next b
  End If
Next a

For a = LBound(CalculationList) To UBound(CalculationList)
  If Not CalculationList(a) = "Done" Then
    For b = LBound(AllCuts) To UBound(AllCuts)
      If AllCuts(b) = "Type" Then
       b = b + 1
       good = False
        Do Until AllCuts(b + 1) = "end"
         b = b + 1
         If CutStatus(b) = "" Then good = True
        Loop
         If good = False Then CalculationList(a) = "Done"
     End If
    Next b
  End If
Next a


Finished = True
For a = 1 To UBound(CalculationList)
 If Not CalculationList(a) = "Done" Then Finished = False
Next a

End Sub

Sub SetCalculationList()
Dim a As Integer, b As Integer, c As Integer, good As Boolean, good2 As Boolean
Dim checkb As Boolean, ca As Worksheet, ExistingWire As Boolean, arr As Variant

Set ca = ThisWorkbook.Worksheets("Calculate")

b = 2
Do Until ca.Cells(7, b + 1).Value = ""
b = b + 1
a = a + 1
ReDim Preserve CalculationList(1 To b + -2)
    CalculationList(a) = ca.Cells(7, b).Value
Loop
checkb = checkarray(CalculationList)
If checkb = False Then End

For Each arr In CalculationList
 ExistingWire = sCheckWire(arr)
 If ExistingWire = False Then
  MsgBox ("Wire type: " & CalculationList(a) & " is not found, please create it in the settings page")
  End
 End If
Next arr

TypeList = CalculationList

End Sub

Sub CollectCuts()
Dim a As Integer, b As Integer, c As Integer, ca As Worksheet
Set ca = ThisWorkbook.Worksheets("Calculate")

a = 2
Do Until ca.Cells(7, a + 1).Value = ""
a = a + 1
b = 7
c = c + 1
ReDim Preserve AllCuts(1 To c + 1)
AllCuts(c) = "Type"
c = c + 1
AllCuts(c) = ca.Cells(7, a).Value
 Do Until ca.Cells(b + 1, a).Value = ""
    b = b + 1
    c = c + 1
    ReDim Preserve AllCuts(1 To c)
    AllCuts(c) = ca.Cells(b, a).Value
 Loop
 c = c + 1
 ReDim Preserve AllCuts(1 To c)
 AllCuts(c) = "end"
Loop
    
ReDim CutStatus(1 To UBound(AllCuts))
ReDim SpoolCutRange(1 To UBound(AllCuts))

End Sub

Sub CollectPreCuts()
Dim a As Integer, b As Integer, c As Integer, d As Integer
Dim hc As Worksheet, lc As Worksheet

Set hc = ThisWorkbook.Worksheets("HIGH CUT")
Set lc = ThisWorkbook.Worksheets("LOW CUT")

For a = LBound(CalculationList) To UBound(CalculationList)
Do Until hc.Cells(2, b + 1).Value = ""
c = 0
b = b + 1
 If hc.Cells(2, b).Value = CalculationList(a) Then
ReDim Preserve HighCuts(1 To d + 2)
d = d + 1
HighCuts(d) = "Type"
d = d + 1
HighCuts(d) = hc.Cells(2, b).Value
c = 2
  Do Until hc.Cells(c + 1, b).Value = ""
    c = c + 1
    d = d + 1
    ReDim Preserve HighCuts(1 To d)
    HighCuts(d) = hc.Cells(c, b).Value
  Loop
  d = d + 1
  ReDim Preserve HighCuts(1 To d)
  HighCuts(d) = "end"
End If
Loop
Next a
b = 0
d = 0

For a = LBound(CalculationList) To UBound(CalculationList)
Do Until lc.Cells(2, b + 1).Value = ""
c = 1
b = b + 1
 If lc.Cells(2, b).Value = CalculationList(a) Then
 ReDim Preserve LowCuts(1 To d + 2)
 d = d + 1
 LowCuts(d) = "Type"
 d = d + 1
 LowCuts(d) = lc.Cells(2, b).Value
 c = 2
   Do Until lc.Cells(c + 1, b).Value = ""
     c = c + 1
     d = d + 1
     ReDim Preserve LowCuts(1 To d)
     LowCuts(d) = lc.Cells(c, b).Value
   Loop
   d = d + 1
   ReDim Preserve LowCuts(1 To d)
   LowCuts(d) = "end"
End If
Loop
Next a

ConLowCuts = LowCuts
ConHighCuts = HighCuts

End Sub

Sub FindPreCuts()
Dim a As Integer, b As Integer, c As Integer, CurrentA As Integer, CurrentB As Integer, d As Integer
Dim hc As Worksheet, lc As Worksheet, bk As Worksheet

Set hc = ThisWorkbook.Worksheets("HIGH CUT")
Set lc = ThisWorkbook.Worksheets("LOW CUT")
Set bk = ThisWorkbook.Worksheets("BULK")

For a = LBound(AllCuts) To UBound(AllCuts)
 If AllCuts(a) = "Type" Then
 CurrentA = a
   For b = LBound(HighCuts) To UBound(HighCuts)
     If HighCuts(b) = "Type" Then
       If HighCuts(b + 1) = AllCuts(a + 1) Then
      a = a + 1
      CurrentB = b
       Do Until AllCuts(a + 1) = "end"
        a = a + 1
        b = CurrentB + 1
         Do Until HighCuts(b + 1) = "end"
           b = b + 1
          If CutStatus(a) = "" Then
           If HighCuts(b) = AllCuts(a) Then
             CutStatus(a) = "Pre-Cut"
             HighCuts(b) = "Selected"
           End If
          End If
         Loop
        Loop
       End If
     End If
    Next b
   For b = LBound(LowCuts) To UBound(LowCuts)
   a = CurrentA
     If LowCuts(b) = "Type" Then
       If LowCuts(b + 1) = AllCuts(a + 1) Then
     a = a + 1
     CurrentB = b
       Do Until AllCuts(a + 1) = "end"
        a = a + 1
        b = CurrentB + 1
         Do Until LowCuts(b + 1) = "end"
           b = b + 1
          If CutStatus(a) = "" Then
           If LowCuts(b) = AllCuts(a) Then
             CutStatus(a) = "Pre-Cut"
             LowCuts(b) = "Selected"
           End If
          End If
         Loop
        Loop
       End If
     End If
    Next b
  For b = LBound(SpoolArray) To UBound(SpoolArray)
   a = CurrentA
     If SpoolArray(b) = "Type" Then
        If SpoolArray(b + 1) = AllCuts(a + 1) Then
      a = a + 1
      CurrentB = b + 1
       Do Until AllCuts(a + 1) = "end"
        b = CurrentB
        a = a + 1
        b = b + 2
         Do Until SpoolArray(b + 1) = "end"
           b = b + 1
          If CutStatus(a) = "" Then
           If SpoolArray(b) = AllCuts(a) Then
             CutStatus(a) = "Pre-Cut"
             SpoolArray(b) = "Done"
           End If
          End If
         Loop
        Loop
        End If
     End If
    Next b
 End If
Next a
        
End Sub
Sub CountBulk()
Dim a As Integer, b As Integer, c As Integer, bcnt1 As Integer
Dim bcnt2 As Integer, d As Integer, bk As Worksheet

Set bk = ThisWorkbook.Worksheets("BULK")

a = 1
Do Until bk.Cells(2, a + 1) = ""
a = a + 1
d = d + 2
ReDim Preserve SpoolQTY(1 To d)
bcnt1 = 0
bcnt2 = 1
 For b = 3 To 100
    If b = 3 And bk.Cells(b, a).Value = "" Then bcnt2 = 2
    If bk.Cells(b, a).Value > "" Then bcnt1 = bcnt1 + 1
 Next b
 For b = 3 To 3 + (bcnt1 + -1)
    If bk.Cells(b, a).Value = "" Then bcnt2 = 2
 Next b
 c = c + 1
 SpoolQTY(c) = bcnt1
 c = c + 1
 SpoolQTY(c) = bcnt2
Loop
       
End Sub

Sub AdjustBulk()
Dim a As Integer, b As Integer, c As Integer, Spools() As Integer
Dim e As Integer, d As Integer, F As Integer, bk As Worksheet

Set bk = ThisWorkbook.Worksheets("BULK")
a = 1
Do Until bk.Cells(2, a + 1).Value = ""
a = a + 1
If F > 0 Then F = F + 2
If F = 0 Then F = 1
b = 2
c = 0
If SpoolQTY(F) > 0 Then ReDim Spools(1 To SpoolQTY(F))
If SpoolQTY(F) = 0 Then GoTo Line1
If SpoolQTY(F + 1) = 1 Then GoTo Line1
  Do Until c = SpoolQTY(F)
  b = b + 1
   If bk.Cells(b, a).Value > "" Then
      c = c + 1
      Spools(c) = bk.Cells(b, a).Value
      bk.Cells(b, a).Value = ""
   End If
  Loop
   e = 0
 For d = 3 To c + 2
   e = e + 1
   bk.Cells(d, a).Value = Spools(e)
 Next d
Line1:
Loop
  
End Sub

Sub SetSpoolArray()
Dim a As Integer, b As Integer, c As Integer, d As Integer
Dim cal As Integer, e As Integer, bk As Worksheet

Set bk = ThisWorkbook.Worksheets("BULK")

ReDim ActiveSpoolRange(1 To UBound(CalculationList))

For cal = 1 To UBound(CalculationList)
a = 1
Do Until bk.Cells(2, a + 1).Value = ""
a = a + 1
  If bk.Cells(2, a).Value = CalculationList(cal) Then
b = 2
  Do Until bk.Cells(b + 1, a).Value = ""
      b = b + 1
  Loop
     If b = 2 Then b = 3
  On Error GoTo Line1
    c = UBound(SpoolArray) + 1
Line1:
    If c = 0 Then c = 1
  ReDim Preserve SpoolArray(1 To (b + -2) + c + 5)
        SpoolArray(c) = "Type"
    c = c + 1
        SpoolArray(c) = bk.Cells(2, a).Value
    c = c + 1
        SpoolArray(c) = "Col"
    c = c + 1
        SpoolArray(c) = a
    c = c + 1
        SpoolArray(c) = "BulkSpools"
    c = c + 1
        e = e + 1
        ActiveSpoolRange(e) = c
    For d = 3 To b
        If bk.Cells(d, a).Value > "" Then SpoolArray(c) = bk.Cells(d, a).Value
        If bk.Cells(d, a).Value = "" Then SpoolArray(c) = "Empty"
    c = c + 1
    Next d
        SpoolArray(c) = "end"
End If
Loop
Next cal
        ReDim SpoolStatus(1 To UBound(SpoolArray))
        SpoolResults = SpoolArray
                  Debug.Print UBound(SpoolResults)
End Sub

Sub SetActiveSpool()
Dim a As Integer, b As Integer, bk As Worksheet
Set bk = ThisWorkbook.Worksheets("BULK")

ReDim ActiveSpool(1 To UBound(CalculationList))

For a = 1 To UBound(CalculationList)
b = 1
 Do Until CalculationList(a) = bk.Cells(2, b).Value
  b = b + 1
 Loop
  ActiveSpool(a) = bk.Cells(3, b).Value
Next a

End Sub

Sub align()
Dim a As Integer, cnt1 As Integer

aa = aa + 1
If aa > UBound(CalculationList) Then
  aa = 1
  Lap = Lap + 1
End If

If Lap = 0 Then Lap = 1

Do Until CalculationList(aa) <> "Done" Or cnt1 > UBound(CalculationList)
    If CalculationList(aa) = "Done" Then
       aa = aa + 1
       If aa > UBound(CalculationList) Then
        aa = 1
        If Not CalculationList(aa) = "Done" Then Lap = Lap + 1
       End If
    End If
    cnt1 = cnt1 + 1
Loop

For a = LBound(AllCuts) To UBound(AllCuts)
   If AllCuts(a) = CalculationList(aa) Then bb = a
Next a
  
For a = LBound(SpoolArray) To UBound(SpoolArray)
    If SpoolArray(a) = CalculationList(aa) Then Exit For
Next a

End Sub

Sub identify()
Dim a As Integer, arr As Variant
Dim sum As Integer, checkb As Boolean

LowQTY = False

a = ActiveSpoolRange(aa)
Do Until SpoolArray(a + 1) = "end"
  a = a + 1
  sum = sum + SpoolArray(a)
Loop
If sum <= 3000 Then LowQTY = True
        
Call FindSaveRanges(CalculationList(aa))
CalculationStart = FindCalcStart(CalculationList(aa))
PrefMaxTrim = FindTrim(CalculationList(aa))
grouping = FindGrouping(CalculationList(aa))
Mode = FindMode(CalculationList(aa))
ManualLaunch = FindManual(CalculationList(aa))
SpecificCuts = FindSpec(CalculationList(aa))
BaseIncrements = FindBase(CalculationList(aa))
qThresholds = FindCurrentThresholds(CalculationList(aa), "ToThresh")
qMaxThresholds = FindCurrentThresholds(CalculationList(aa), "ToMax")
ThresholdStages = FindCurrentThresholds(CalculationList(aa), "Stages")

checkb = checkarray(ManualLaunch)
If checkb = True Then
 For Each arr In ManualLaunch
  If arr = "Calculation" Then StartOnCalc(aa) = True
  If arr = "Start" Then StartOnStart(aa) = True
  If arr = "Trim" Then StartOnTrim(aa) = True
  If arr = "Cut" Then StartOnCut(aa) = True
 Next arr
End If

mStandard = False
mFree = False
mCritical = False

If Mode = "Standard" Then mStandard = True
If Mode = "Free" Then mFree = True
If Mode = "Critical" Then mCritical = True

End Sub

Sub DefineCuts()

Dim a As Integer, b As Integer, c As Integer, CurrentCutsRange As Integer, CurrentLastRange As Integer, arr As Variant
Dim ca As Worksheet
Set ca = ThisWorkbook.Worksheets("calculate")
TotalSum = 0

a = 1
Do Until AllCuts(a) = CalculationList(aa)
 a = a + 1
Loop
        CurrentCutsRange = a
Do Until AllCuts(a) = "end"
a = a + 1
Loop
        CurrentLastRange = a


For a = CurrentCutsRange + 1 To CurrentLastRange + -1
     If Not CutStatus(a) = "Cut" And Not CutStatus(a) = "Pre-Cut" Then
      TotalSum = TotalSum + AllCuts(a)
      c = c + 1
      ReDim Preserve Cutslist(1 To c)
      Cutslist(c) = AllCuts(a)
    End If
Next a

                                                        AllCutCombinations = FindAllCombinations(Cutslist, 0, 0)
      
                                                                                                                
End Sub

Sub NewSpool()

Dim a As Integer, b As Integer, c As Integer, good As Boolean, Hspool As Integer, good2 As Boolean, Ds As Integer

c = UBound(DeSpool)
ReDim Preserve DeSpool(1 To c + 2)

good = False
For a = 1 To UBound(Cutslist)
 If SpoolResults(ActiveSpoolRange(aa)) > Cutslist(a) Then good = True
Next a

If good = False Then
For a = 1 To UBound(SpoolResults)
 If SpoolResults(a) = "Type" Then
   If SpoolResults(a + 1) = CalculationList(aa) Then
    a = a + 4
    Do Until SpoolResults(a + 1) = "end"
     a = a + 1
     If SpoolResults(a) > Hspool Then
      Hspool = SpoolResults(a)
      good2 = True
      b = a
     End If
    Loop
   End If
 End If
Next a

Ds = ActiveSpool(aa)
ActiveSpool(aa) = Hspool
ActiveSpoolRange(aa) = b
End If


c = c + 1
DeSpool(c) = Lap
c = c + 1
If good = False And good2 = True Then DeSpool(c) = Ds
If good = False And good2 = False Then DeSpool(c) = ""
If good = True Then DeSpool(c) = ""

End Sub

Sub FindDifference()
Dim a As Integer, b As Integer, c As Integer
Dim checkb As Boolean

LowSpec = FindLowestNumber(SpecificCuts)

Erase InputDiff
Erase InputDiffQTY

b = 1
For a = LBound(AllCutCombinations) To UBound(AllCutCombinations)
If AllCutCombinations(a) = "Cuts" Then c = a
    If AllCutCombinations(a) = "Sum" Then
       If -AllCutCombinations(a + 1) + ActiveSpool(aa) <= CalculationStart _
       And -AllCutCombinations(a + 1) + ActiveSpool(aa) >= LowSpec Then
         ReDim Preserve InputDiff(1 To b)
         ReDim Preserve InputDiffLoc(1 To b)
         InputDiff(b) = -AllCutCombinations(a + 1) + ActiveSpool(aa)
         InputDiffLoc(b) = c
         b = b + 1
       End If
    End If
Next a

NoSpec = False
checkb = checkarray(InputDiff)
If checkb = False Then NoSpec = True
                                                                                                                                                   
End Sub
                                                            
Sub AltSpec()
Dim a As Integer

ReDim OptimalSpecCombination(1 To 8)
a = a + 1
    OptimalSpecCombination(a) = "Cuts"
a = a + 1
    OptimalSpecCombination(a) = 0
a = a + 1
    OptimalSpecCombination(a) = "Sum"
a = a + 1
    OptimalSpecCombination(a) = 0
a = a + 1
    OptimalSpecCombination(a) = "Difference"
a = a + 1
    OptimalSpecCombination(a) = -TotalSum + ActiveSpool(aa)
a = a + 1
    OptimalSpecCombination(a) = "DiffQTY"
a = a + 1
    OptimalSpecCombination(a) = UBound(Cutslist)
    
End Sub


Sub FindAllMultiples()
Dim a As Long, b As Long, c As Long, d As Integer, HighestDifference As Integer, InputQTY() As Integer, inputmult() As Integer
Dim QTYcount As Integer, Sumbox() As Integer, LowestDifference As Integer, Thresholds() As Integer, MaxThresholds() As Integer, arr As Variant

Call FindSaveRanges(CalculationList(aa))
Thresholds = FindThresh(CalculationList(aa))
MaxThresholds = FindMaxThresh(CalculationList(aa))

HighestDifference = FindHighestNumber(InputDiff)

ReDim InputQTY(1 To UBound(SpecificCuts))

For a = LBound(SpecificCuts) To UBound(SpecificCuts)
   If SpecificCuts(a) <= HighestDifference Then
      InputQTY(a) = HighestDifference \ SpecificCuts(a)
      QTYcount = QTYcount + InputQTY(a)
   End If
   If SpecificCuts(a) > HighestDifference Then InputQTY(a) = 0
Next a

ReDim inputmult(1 To QTYcount)

b = 1
For a = LBound(SpecificCuts) To UBound(SpecificCuts)
    For b = b To InputQTY(a) + b + -1
        inputmult(b) = SpecificCuts(a)
    Next b
Next a


LowestDifference = FindLowestNumber(InputDiff)

ReDim InputDiffQTY(1 To UBound(InputDiff))

For a = LBound(InputDiffLoc) To UBound(InputDiffLoc)
    For b = LBound(AllCutCombinations) To UBound(AllCutCombinations)
        If b = InputDiffLoc(a) Then
           Do Until AllCutCombinations(b) = "end"
              b = b + 1
              If AllCutCombinations(b) = "Number Of Cuts" Then InputDiffQTY(a) = AllCutCombinations(b + 1)
           Loop
        End If
    Next b
Next a


SpecMult = FindAllCombinations(inputmult, HighestDifference, LowestDifference)
OptimalSpecCombination = FindOptimalSpecCombination(SpecMult, InputDiff, PrefMaxTrim, Thresholds, MaxThresholds, SpecificCuts, InputDiffQTY, LowQTY)

For Each arr In OptimalSpecCombination
Debug.Print arr
Next arr
End Sub

Sub Store()
Dim a As Integer, b As Integer, c As Integer, opDiff As Integer, DiffQTY As Integer, opSum As Integer, DiffRange As Integer, LastDiffRange As Integer
Dim CutsRange As Integer, SpecQTY As Integer, Size As Integer, checkb As Boolean, bk As Worksheet


Set bk = ThisWorkbook.Worksheets("BULK")

opDiff = OptimalSpecCombination(UBound(OptimalSpecCombination) + -2)
opSum = -opDiff + ActiveSpool(aa)
DiffQTY = OptimalSpecCombination(UBound(OptimalSpecCombination))

DiffRange = 0
For a = LBound(AllCutCombinations) To UBound(AllCutCombinations)
  If AllCutCombinations(a) = "Cuts" Then b = a
    If AllCutCombinations(a) = "Sum" Then
      If AllCutCombinations(a + 1) = opSum Then
         c = a
         Do Until AllCutCombinations(c) = "end"
            c = c + 1
            If AllCutCombinations(c) = "Number Of Cuts" Then
               If AllCutCombinations(c + 1) = DiffQTY Then DiffRange = b
            End If
         Loop
           If DiffRange > 0 Then LastDiffRange = c
       End If
    End If
Next a

For a = LBound(OptimalSpecCombination) To UBound(OptimalSpecCombination)
   If a = 2 Then
    Do Until OptimalSpecCombination(a) = "Sum"
       SpecQTY = SpecQTY + 1
       a = a + 1
    Loop
   End If
Next a

If opDiff < LowSpec Then SpecQTY = SpecQTY + 1

For a = DiffRange To LastDiffRange
    If AllCutCombinations(a) = "Cuts" Then
       Do Until AllCutCombinations(a) = "Sum"
          a = a + 1
          b = bb
            Do Until AllCuts(b) = "end"
             b = b + 1
               If AllCutCombinations(a) = AllCuts(b) And CutStatus(b) = "" Then
                  CutStatus(b) = "Cut"
                  SpoolCutRange(b) = CurrentSpoolRange
               End If
            Loop
       Loop
    End If
Next a

SpoolResults(ActiveSpoolRange(aa)) = -(opDiff + opSum) + ActiveSpool(aa)
If SpoolResults(ActiveSpoolRange(aa)) = 0 Then SpoolStatus(ActiveSpoolRange(aa)) = "Done"

Size = SpecQTY + DiffQTY + 18
checkb = checkarray(CutResults)
    
     If checkb = True Then b = UBound(CutResults) + 1
     If checkb = False Then b = 1
      
ReDim Preserve CutResults(1 To Size + b)
      CutResults(b) = "Lap"
      b = b + 1
      CutResults(b) = Lap
      b = b + 1
      CutResults(b) = "Type"
      b = b + 1
      CutResults(b) = CalculationList(aa)
      b = b + 1
      CutResults(b) = "Spool"
      b = b + 1
      CutResults(b) = ActiveSpool(aa)
      b = b + 1
      CutResults(b) = "Sum Of Cuts"
      b = b + 1
      CutResults(b) = opSum
      b = b + 1
      CutResults(b) = "SpoolDiff"
      b = b + 1
      CutResults(b) = -(opDiff + opSum) + ActiveSpool(aa)
      b = b + 1
      CutResults(b) = "Number Of Cuts"
      b = b + 1
      CutResults(b) = DiffQTY
      b = b + 1
      CutResults(b) = "Number Of Pre-Cuts"
      b = b + 1
      CutResults(b) = SpecQTY
      
    
For a = DiffRange To LastDiffRange
 If AllCutCombinations(a) = "Cuts" Then
      b = b + 1
      CutResults(b) = "Requested Cuts"
   Do Until AllCutCombinations(a + 1) = "Sum"
      a = a + 1
      b = b + 1
      CutResults(b) = AllCutCombinations(a)
   Loop
 End If
Next a
   
      b = b + 1
      CutResults(b) = "Pre-Cuts"
      
a = 1
Do Until OptimalSpecCombination(a + 1) = "Sum"
 b = b + 1
 a = a + 1
   CutResults(b) = OptimalSpecCombination(a)
Loop


    
    If opDiff < LowSpec Then
      b = b + 1
      CutResults(b) = OptimalSpecCombination(a)
    End If

     b = b + 1
     CutResults(b) = "end"
             
End Sub

Sub ClearProjection()
Dim a As Integer, b As Integer, c As Integer, good As Boolean, pj As Worksheet
Dim LastCol As Integer, LastRow As Integer, First As Integer
Set pj = ThisWorkbook.Worksheets("Projection")

a = 1
good = False
Do Until good = True
good = True
  Do Until pj.Cells(1, a).Value = "" And pj.Cells(1, a).Interior.ColorIndex = -4142
  a = a + 1
  Loop
c = 0
First = a
  Do Until c = 10
   a = a + 1
   c = c + 1
   If pj.Cells(1, a).Value > "" Or Not pj.Cells(1, a).Interior.ColorIndex = -4142 Then
    good = False
    Exit Do
   End If
  Loop
If good = True Then a = First
Loop
        LastCol = a

b = 1
good = False
Do Until good = True
good = True
  Do Until pj.Cells(b, 1).Value = "" And pj.Cells(b, 1).Interior.ColorIndex = -4142
   b = b + 1
  Loop
c = 0
First = b
  Do Until c = 10
   b = b + 1
   c = c + 1
   If pj.Cells(b, 1).Value > "" Or Not pj.Cells(b, 1).Interior.ColorIndex = -4142 Then
    good = False
    Exit Do
   End If
  Loop
   If good = True Then b = First
Loop
        LastRow = b

Range(pj.Cells(1, 1), pj.Cells(LastRow, LastCol)).Delete

End Sub

Sub CheckPreCuts()
Dim checkb As Boolean

checkb = checkarray(CutResults)
If checkb = False Then OnlyPreCuts = True
End Sub

Sub ProjectOnlyPreCuts()
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, pj As Worksheet
Dim HighestPreCut As Integer, HighestPending As Integer
Dim FirstPreCutRange As Integer, LastPreCutRange As Integer
Dim FirstPendingRange As Integer, LastPendingRange As Integer
Set pj = ThisWorkbook.Worksheets("Projection")

HighestPreCut = FindHighestCut(AllCuts, CutStatus, "Pre-Cut")
HighestPending = FindHighestCut(AllCuts, CutStatus, "")

FirstPreCutRange = 3
If HighestPreCut > 0 Then LastPreCutRange = FirstPreCutRange + HighestPreCut + -1
If HighestPreCut = 0 Then LastPreCutRange = FirstPreCutRange
FirstPendingRange = LastPreCutRange + 1
If HighestPending > 0 Then LastPendingRange = FirstPendingRange + HighestPending + -1
If HighestPending = 0 Then LastPendingRange = FirstPendingRange
 
Call ProjectColors(1, 1, 119, 119, 119, 28, True, "TOTALS", 45, True, UBound(TypeList) + 1)
Call ProjectColors(2, 2, 0, 112, 192, 14, True, "Wire Type", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstPreCutRange, LastPreCutRange, 146, 208, 80, 14, False, "From Pre-Cuts", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstPendingRange, LastPendingRange, 255, 80, 80, 14, False, "Pending", 30, False, UBound(TypeList) + 1)

For a = 1 To UBound(TypeList)
 pj.Cells(2, a + 1).Value = TypeList(a)
Next a
    
For a = LBound(AllCuts) To UBound(AllCuts)
 If AllCuts(a) = "Type" Then
   a = a + 1
   b = 1
   c = FirstPreCutRange
   e = FirstPendingRange
    Do Until pj.Cells(2, b).Value = AllCuts(a)
     b = b + 1
    Loop
    Do Until AllCuts(a + 1) = "end"
     a = a + 1
      If CutStatus(a) = "Pre-Cut" Then
         pj.Cells(c, b).Value = AllCuts(a)
       c = c + 1
      End If
      If CutStatus(a) = "" Then
         pj.Cells(e, b).Value = AllCuts(a)
       e = e + 1
      End If
    Loop
  End If
Next a
 
 
End Sub

Sub FindProjectionRanges()
Dim a As Integer, b As Integer, c As Integer, CutRange As Integer, SpecRange As Integer, PrevRange As Integer

Lap = 0
For a = 1 To UBound(CutResults)
 If CutResults(a) = "Lap" Then
  If Lap < CutResults(a + 1) Then Lap = CutResults(a + 1)
 End If
Next a
   
ReDim ProjectionRanges(1 To Lap)

For a = 1 To Lap
  For b = LBound(CutResults) To UBound(CutResults)
     If CutResults(b) = "Lap" Then
       If CutResults(b + 1) = a Then
          Do Until CutResults(b) = "Number Of Cuts"
            b = b + 1
          Loop
           CutRange = CutResults(b + 1)
          Do Until CutResults(b) = "Number Of Pre-Cuts"
           b = b + 1
          Loop
           SpecRange = CutResults(b + 1)
           If ProjectionRanges(a) + 5 < SpecRange + CutRange + 5 + PrevRange Then ProjectionRanges(a) = SpecRange + CutRange + 5 + PrevRange
       End If
     End If
  Next b
PrevRange = ProjectionRanges(a)
Next a

End Sub


Sub ProjectResults1()
Dim a As Integer, b As Integer, c As Integer, FirstRange As Integer, d As Integer
Dim pj As Worksheet, More As Boolean

Dim CurrentNOC As Integer, CurrentNOPC As Integer
Dim FirstC As Integer, LastC As Integer, FirstAC As Integer, LastAC As Integer
Dim red1 As Integer, green1 As Integer, blue1 As Integer
Dim red2 As Integer, green2 As Integer, blue2 As Integer
Dim red3 As Integer, green3 As Integer, blue3 As Integer
Dim red4 As Integer, green4 As Integer, blue4 As Integer
Dim red5 As Integer, green5 As Integer, blue5 As Integer
Dim red6 As Integer, green6 As Integer, blue6 As Integer
Dim red7 As Integer, green7 As Integer, blue7 As Integer

red1 = 119
green1 = 119
blue1 = 119

red2 = 0
green2 = 112
blue2 = 192

red3 = 248
green3 = 203
blue3 = 173

red4 = 208
green4 = 206
blue4 = 206

red5 = 189
green5 = 215
blue5 = 238

red6 = 255
green6 = 255
blue6 = 255

red7 = 153
green7 = 255
blue7 = 153

Set pj = ThisWorkbook.Worksheets("Projection")

pj.Columns("A").ColumnWidth = 18.5

For a = 1 To Lap

If FirstRange > 0 Then FirstRange = ProjectionRanges(a + -1) + 1
If FirstRange = 0 Then FirstRange = 1

  For b = LBound(TypeList) + 1 To UBound(TypeList) + 1

CurrentNOC = FindCutRanges(CutResults, a, "Number Of Cuts")
CurrentNOPC = FindCutRanges(CutResults, a, "Number Of Pre-Cuts")
    
FirstC = FirstRange + 5
If CurrentNOC > 0 Then LastC = FirstC + CurrentNOC + -1
If CurrentNOC = 0 Then LastC = FirstC
FirstAC = LastC + 1
If CurrentNOPC > 0 Then LastAC = FirstAC + CurrentNOPC + -1
If CurrentNOPC = 0 Then LastAC = FirstAC

Call ProjectColors(FirstRange, FirstRange, red1, green1, blue1, 14, True, "LAP " & a, 30, True, UBound(TypeList) + 1)
Call ProjectColors(FirstRange + 1, FirstRange + 1, red2, green2, blue2, 14, True, "Wire Type", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstRange + 2, FirstRange + 2, red3, green3, blue3, 14, True, "Initial Spool", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstRange + 3, FirstRange + 3, red4, green4, blue4, 14, True, "Sum Of Cuts", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstRange + 4, FirstRange + 4, red5, green5, blue5, 14, True, "Resulting Spool", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstC, LastC, red6, green6, blue6, 14, True, "Requested Cuts", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstAC, LastAC, red7, green7, blue7, 14, True, "Additional Cuts", 30, False, UBound(TypeList) + 1)

pj.Cells(FirstRange + 1, b).Value = TypeList(b + -1)

      For c = LBound(CutResults) To UBound(CutResults)
         If CutResults(c) = "Lap" Then
           If CutResults(c + 1) = a And CutResults(c + 3) = TypeList(b + -1) Then
             Do Until CutResults(c) = "end"
              c = c + 1
              If CutResults(c) = "Spool" Then pj.Cells(FirstRange + 2, b).Value = CutResults(c + 1)
              If CutResults(c) = "Sum Of Cuts" Then pj.Cells(FirstRange + 3, b).Value = CutResults(c + 1)
              If CutResults(c) = "SpoolDiff" Then pj.Cells(FirstRange + 4, b).Value = CutResults(c + 1)
              
              If CutResults(c) = "Requested Cuts" Then
               d = FirstRange + 4
               Do Until CutResults(c + 1) = "Pre-Cuts"
                c = c + 1
                d = d + 1
                With pj.Cells(d, b)
                .Value = CutResults(c)
                .Interior.Color = RGB(red6, green6, blue6)
                .Font.Size = 12
                .Font.ColorIndex = 1
                .Font.Bold = False
                End With
               Loop
              End If
           
              
              If CutResults(c) = "Pre-Cuts" Then
                Do Until CutResults(c + 1) = "end"
                 c = c + 1
                 d = d + 1
                 With pj.Cells(d, b)
                 .Value = CutResults(c)
                 .Interior.Color = RGB(red7, green7, blue7)
                 .Font.Size = 12
                 .Font.ColorIndex = 1
                 .Font.Bold = False
                 End With
                Loop
             End If
           
            Loop
           End If
        End If
     Next c
  Next b
Next a

End Sub

Sub ProjectResults2()
Dim a As Integer, b As Integer, c As Integer, CurrentType As Integer, pj As Worksheet, d As Integer, e As Integer
Dim FirstAddRange As Integer, LastAddRange As Integer, FirstPreCutRange As Integer, LastPreCutRange As Integer, FirstPendingRange As Integer
Dim LastPendingRange As Integer, FirstCutRange As Integer, LastCutRange As Integer, AddCuts() As Variant, HighestAdd As Integer
Dim TotalsRange As Integer, TypeRange As Integer, HighestPreCut As Integer, HighestCut As Integer, HighestPending As Integer
Dim FirstEmpty As Integer, LastEmpty As Integer
Dim red0 As Integer, green0 As Integer, blue0 As Integer
Dim red1 As Integer, green1 As Integer, blue1 As Integer
Dim red2 As Integer, green2 As Integer, blue2 As Integer
Dim red3 As Integer, green3 As Integer, blue3 As Integer
Dim red4 As Integer, green4 As Integer, blue4 As Integer
Dim red5 As Integer, green5 As Integer, blue5 As Integer

AddCuts = FindAdditionalCuts(TypeList, CutResults)
HighestPreCut = FindHighestCut(AllCuts, CutStatus, "Pre-Cut")
HighestCut = FindHighestCut(AllCuts, CutStatus, "Cut")
HighestPending = FindHighestCut(AllCuts, CutStatus, "")
HighestAdd = FindHighestAdd(AddCuts)

TotalsRange = ProjectionRanges(UBound(ProjectionRanges)) + 1
TypeRange = ProjectionRanges(UBound(ProjectionRanges)) + 2
FirstPreCutRange = TypeRange + 1
If HighestPreCut > 0 Then LastPreCutRange = FirstPreCutRange + HighestPreCut + -1
If HighestPreCut = 0 Then LastPreCutRange = FirstPreCutRange
FirstCutRange = LastPreCutRange + 1
If HighestCut > 0 Then LastCutRange = FirstCutRange + HighestCut + -1
If HighestCut = 0 Then LastCutRange = FirstCutRange
FirstPendingRange = LastCutRange + 1
If HighestPending > 0 Then LastPendingRange = FirstPendingRange + HighestPending + -1
If HighestPending = 0 Then LastPendingRange = FirstPendingRange
FirstAddRange = FirstPendingRange + 1
If HighestAdd > 0 Then LastAddRange = FirstAddRange + HighestAdd + -1
If HighestAdd = 0 Then LastAddRange = FirstAddRange

Set pj = ThisWorkbook.Worksheets("Projection")

red0 = 119
green0 = 119
blue0 = 119

red1 = 0
green1 = 112
blue1 = 192

red2 = 146
green2 = 208
blue2 = 80

red3 = 0
green3 = 176
blue3 = 80

red4 = 255
green4 = 80
blue4 = 80

red5 = 153
green5 = 255
blue5 = 153

Call ProjectColors(TotalsRange, TotalsRange, red0, green0, blue0, 28, True, "TOTALS", 45, True, UBound(TypeList) + 1)
Call ProjectColors(TypeRange, TypeRange, red1, green1, blue1, 14, True, "Wire Type", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstPreCutRange, LastPreCutRange, red2, green2, blue2, 14, True, "From Pre-Cuts", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstCutRange, LastCutRange, red3, green3, blue3, 14, True, "From Spool", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstPendingRange, LastPendingRange, red4, green4, blue4, 14, True, "Pending", 30, False, UBound(TypeList) + 1)
Call ProjectColors(FirstAddRange, LastAddRange, red5, green5, blue5, 14, True, "Additional Cuts", 30, False, UBound(TypeList) + 1)

b = 1
For a = 1 To UBound(TypeList)
b = b + 1
 pj.Cells(TypeRange, b).Value = TypeList(a)
Next a

For a = LBound(AllCuts) To UBound(AllCuts)
 If AllCuts(a) = "Type" Then
   a = a + 1
   b = 1
   c = FirstPreCutRange
   d = FirstCutRange
   e = FirstPendingRange
    Do Until pj.Cells(2, b).Value = AllCuts(a)
     b = b + 1
    Loop
    Do Until AllCuts(a + 1) = "end"
     a = a + 1
      If CutStatus(a) = "Pre-Cut" Then
         pj.Cells(c, b).Value = AllCuts(a)
       c = c + 1
      End If
      If CutStatus(a) = "Cut" Then
         pj.Cells(d, b).Value = AllCuts(a)
       d = d + 1
      End If
      If CutStatus(a) = "" Then
         pj.Cells(d, b).Value = AllCuts(a)
       e = e + 1
      End If
    Loop
  End If
Next a

b = 1
For a = LBound(AddCuts) To UBound(AddCuts)
 If AddCuts(a) = "Type" Then
  a = a + 1
  c = FirstAddRange
  b = b + 1
  Do Until AddCuts(a + 1) = "end"
   a = a + 1
   pj.Cells(c, b).Value = AddCuts(a)
   c = c + 1
  Loop
 End If
Next a
   
End Sub

Sub Manual()
Dim a As Integer, b As Integer, c As Integer, MN As Worksheet
Dim red1 As Integer, green1 As Integer, blue1 As Integer
Dim red2 As Integer, green2 As Integer, blue2 As Integer
Dim red3 As Integer, green3 As Integer, blue3 As Integer
Dim red4 As Integer, green4 As Integer, blue4 As Integer
Dim red5 As Integer, green5 As Integer, blue5 As Integer
Dim Cuts() As Integer, Spools() As Integer, PreCuts() As Integer
Dim FirstCutsRange As Integer, FirstSpoolsRange As Integer, FirstPreCutsRange As Integer
Dim LastCutsRange As Integer, LastSpoolsRange As Integer, LastPreCutsRange As Integer
Dim Header1 As Integer, Header2 As Integer, Header3 As Integer, HighestCount As Integer, Header4 As Integer, Header5 As Integer
Dim CutsCount As Integer, SpoolsCount As Integer, PreCutsCount As Integer, CountDiv As Integer, CountMult As Integer
Dim FirstBorderRange As Integer, SecondBorderRange As Integer, ThirdBorderRange As Integer, FourthBorderRange
Dim FirstSelectedRange As Integer, LastSelectedRange As Integer, FirstSelectedSpoolRange As Integer, LastSelectedSpoolRange As Integer
Dim FirstSelectedPreCutRange As Integer, LastSelectedPreCutRange As Integer, done As Boolean, CurrentSelectedCut As Integer, CurrentCutsCol As Integer
Dim FirstIncRange As Integer, Header0 As Integer

If StartOnStart(aa) = True Then ReDim CutStatus(1 To UBound(AllCuts))

Set MN = ThisWorkbook.Worksheets("Manual")
red1 = 0
green1 = 32
blue1 = 96

red2 = 255
green2 = 255
blue2 = 255

red3 = 0
green3 = 176
blue3 = 240

red4 = 153
green4 = 255
blue4 = 204

red5 = 155
green5 = 194
blue5 = 230


Cuts = GetCuts(AllCuts, CutStatus, CalculationList(aa))
Spools = GetSpools(CalculationList(aa), SpoolResults)
PreCuts = GetPreCuts(ConHighCuts, ConLowCuts, CalculationList(aa))

If UBound(Cuts) >= 6 Then CutsCount = UBound(Cuts) \ 6
If UBound(Cuts) < 6 Then CutsCount = 0
If UBound(Spools) >= 6 Then SpoolsCount = UBound(Spools) \ 6
If UBound(Spools) < 6 Then SpoolsCount = 0
If UBound(PreCuts) >= 6 Then PreCutsCount = UBound(PreCuts) \ 6
If UBound(PreCuts) < 6 Then PreCutsCount = 0

Header0 = 1
Header1 = 2
FirstCutsRange = Header1 + 1
LastCutsRange = FirstCutsRange + CutsCount
Header2 = LastCutsRange + 1
FirstSpoolsRange = Header2 + 1
LastSpoolsRange = FirstSpoolsRange + SpoolsCount
Header3 = LastSpoolsRange + 1
FirstPreCutsRange = Header3 + 1
LastPreCutsRange = FirstPreCutsRange + PreCutsCount

If HighestCount < UBound(Cuts) Then HighestCount = UBound(Cuts)
If HighestCount < UBound(Spools) Then HighestCount = UBound(Spools)
If HighestCount < UBound(PreCuts) Then HighestCount = UBound(PreCuts)
If HighestCount >= 6 Then HighestCount = 6

FirstBorderRange = HighestCount + 1
FirstSelectedRange = FirstBorderRange + 1
FirstSelectedSpoolRange = FirstBorderRange + 1
LastSelectedSpoolRange = FirstSelectedSpoolRange + 3
SecondBorderRange = LastSelectedSpoolRange + 1
FirstSelectedPreCutRange = SecondBorderRange + 1
LastSelectedPreCutRange = FirstSelectedPreCutRange + 2
LastSelectedRange = FirstSelectedPreCutRange + 1
ThirdBorderRange = LastSelectedPreCutRange + 1
CurrentCutsCol = ThirdBorderRange + 1

Header4 = 1
FirstIncRange = Header4 + 3

Call ClearManual
Call ManualColors(Header0, Header0, red1, green1, blue1, 28, True, TypeList(aa), 30, True, 6, 1, True)
Call ManualColors(Header1, Header1, red1, green1, blue1, 28, True, "Requested Cuts", 30, True, HighestCount, 1, True)
Call ManualColors(FirstCutsRange, LastCutsRange, red2, green2, blue2, 14, False, "", 30, False, HighestCount, 1, False)
Call ManualColors(Header2, Header2, red1, green1, blue1, 28, True, "Spools", 30, True, HighestCount, 1, True)
Call ManualColors(FirstSpoolsRange, LastSpoolsRange, red2, green2, blue2, 14, False, "", 30, False, HighestCount, 1, False)
Call ManualColors(Header3, Header3, red1, green1, blue1, 28, True, "Pre-Cuts", 30, True, HighestCount, 1, True)
Call ManualColors(FirstPreCutsRange, LastPreCutsRange, red2, green2, blue2, 14, False, "", 30, False, HighestCount, 1, False)
Call ManualBorder(FirstBorderRange, LastPreCutsRange, 3, 1)

Call ManualColors(1, 1, red1, green1, blue1, 18, True, "Selected Spools", 30, True, LastSelectedRange, FirstSelectedRange, True)
Call ManualColors(1, 1, red1, green1, blue1, 18, True, "Selected Pre-Cuts", 30, True, LastSelectedPreCutRange, FirstSelectedPreCutRange, True)
Call ManualColors(2, 2, red1, green1, blue1, 14, True, "Active", 30, True, FirstSelectedSpoolRange + 1, FirstSelectedSpoolRange, True)
Call ManualColors(2, 2, red1, green1, blue1, 14, True, "Inactive", 30, True, LastSelectedSpoolRange, FirstSelectedSpoolRange + 2, True)
Call ManualColors(3, 3, red3, green3, blue3, 14, False, "", 30, False, LastSelectedSpoolRange, FirstSelectedSpoolRange, False)
Call ManualColors(4, 4, red5, green5, blue5, 14, False, "", 30, False, LastSelectedSpoolRange, FirstSelectedSpoolRange, False)
Call ManualColors(5, 9, red2, green2, blue2, 14, False, "", 30, False, LastSelectedSpoolRange, FirstSelectedSpoolRange, False)
Call ManualBorder(FirstSelectedSpoolRange + 1, LastPreCutsRange, 1, 2)
Call ManualBorder(SecondBorderRange, LastPreCutsRange, 1, 1)

Call ManualColors(2, 9, red2, green2, blue2, 14, False, "", 30, False, LastSelectedPreCutRange, FirstSelectedPreCutRange, False)
Call ManualBorder(ThirdBorderRange, LastPreCutsRange, 3, 1)

Call ManualColors(Header0, Header0, red1, green1, blue1, 28, True, "Actions", 30, True, CurrentCutsCol + 6, CurrentCutsCol, True)
Call AddButton(2, CurrentCutsCol, CurrentCutsCol + 1, "DoneButton", "Done", "Manual")
Call AddButton(3, CurrentCutsCol, CurrentCutsCol + 1, "TransferButton", "Transfer Tool", "Manual")
Call AddButton(4, CurrentCutsCol, CurrentCutsCol + 1, "IncrementButton", "Increment Tool", "Manual")



b = FirstCutsRange
c = 0
For a = 1 To UBound(Cuts)
 c = c + 1
 MN.Cells(b, c).Value = Cuts(a)
 If c = 6 Then
  b = b + 1
  c = 0
 End If
Next a

c = 0
b = FirstSpoolsRange
For a = 1 To UBound(Spools)
 c = c + 1
 MN.Cells(b, c).Value = Spools(a)
 If c = 6 Then
  b = b + 1
  c = 0
 End If
Next a

c = 0
b = FirstPreCutsRange
For a = 1 To UBound(PreCuts)
 c = c + 1
 MN.Cells(b, c).Value = PreCuts(a)
 If c = 6 Then
  b = b + 1
  c = 0
 End If
Next a

Call HighlightWire(AllCuts, CutStatus, SpoolCutRange, TypeList(aa), FirstCutsRange, LastCutsRange, FirstSelectedSpoolRange, FirstSelectedPreCutRange, SpoolResults, LastSelectedPreCutRange, FirstPreCutsRange)
Call CheckSelected(FirstSelectedRange, LastSelectedRange, FirstBorderRange, SecondBorderRange, ThirdBorderRange, FirstSelectedSpoolRange, _
LastSelectedSpoolRange, FirstSelectedPreCutRange, LastSelectedPreCutRange, red2, green2, blue2)

Do Until done = True
DoEvents
If MN.Cells(10, 1).Value = True Then done = True
Loop
                         
End Sub

Function FindDeSpool(Lap)
Dim a As Integer, b As Integer
Dim oValues() As Integer

For a = 1 To UBound(DeSpool)
 If DeSpool(a) = Lap Then
  Do Until a + 1 > UBound(DeSpool) Or DeSpool(a + 1) = Lap + 1
   a = a + 1
   b = b + 1
   ReDim Preserve oValues(1 To b)
   oValues(b) = DeSpool(a)
  Loop
 End If
Next a

FindDeSpool = oValues

End Function

