Attribute VB_Name = "mInventory"
Option Explicit
Dim HighCuts As Range, LowCuts As Range, Bulk As Range
Dim hc As Worksheet, lc As Worksheet, bk As Worksheet

Sub SetWorksheets(worksheetrange, worksheetname)
Set worksheetrange = ThisWorkbook.Worksheets(worksheetname)
End Sub
Sub FindRanges(wirename)
Call SetWorksheets(hc, "HIGH CUT")
Call SetWorksheets(lc, "LOW CUT")
Call SetWorksheets(bk, "BULK")
Call SetRange(wirename, HighCuts, hc, False)
Call SetRange(wirename, LowCuts, lc, False)
Call SetRange(wirename, Bulk, bk, False)
Call CheckWire(wirename)
End Sub

Sub SetRange(wirename, RangeName, Worksheet, top)
Dim a As Integer, b As Integer, FirstRow As Integer

If top = True Then FirstRow = 2
If top = False Then FirstRow = 3

a = 1
Do Until Worksheet.Cells(2, a).Value = wirename
 a = a + 1
Loop

b = 2
Do Until Worksheet.Cells(b + 1, a).Value = ""
 b = b + 1
Loop

If b = 2 Then b = 3

Set RangeName = Range(Worksheet.Cells(FirstRow, a), Worksheet.Cells(b, a))

End Sub

Function FindCurrentThresholds(wirename, mType)
Dim Thresholds() As Integer, MaxThresholds() As Integer, arr1 As Variant, arr2 As Variant, ToMax() As Integer, ToThresh() As Integer
Dim SpecificCuts() As Integer, grouping As Integer, SpecCount() As Integer, Stages() As Variant, a As Integer

Call FindSaveRanges(wirename)
Thresholds = FindThresh(wirename)
MaxThresholds = FindMaxThresh(wirename)
SpecificCuts = FindSpec(wirename)
grouping = FindGrouping(wirename)

Call FindRanges(wirename)
SpecCount = CountMatchingRanges(SpecificCuts, HighCuts, grouping)

For Each arr1 In SpecCount
 For Each arr2 In Thresholds
  a = a + 1
  ReDim Preserve Stages(1 To a)
  ReDim Preserve ToThresh(1 To a)
  ReDim Preserve ToMax(1 To a)
  Stages(a) = (arr1 * arr2) \ arr2
  ToThresh(a) = -arr1 + arr2
 Next arr2
 For Each arr2 In MaxThresholds
  If arr1 >= arr2 Then Stages(a) = "Max"
  ToMax(a) = -arr1 + arr2
 Next arr2
Next arr1
 
If mType = "Stages" Then FindCurrentThresholds = Stages
If mType = "ToMax" Then FindCurrentThresholds = ToMax
If mType = "ToThresh" Then FindCurrentThresholds = ToThresh
End Function

Function FindAllDiscrepancies(mList, wirename)
Dim cell As Range, arr As Variant, Match As Boolean, SpecCuts() As Integer, a As Integer, b As Integer
Dim MatchList() As Integer, grouping As Integer, oBulk() As Integer, oHighCuts() As Integer, oLowCuts() As Integer, check As Boolean
Dim oValues() As Integer, dBulk() As Integer, dHighCuts() As Integer, dLowCuts() As Integer, iType(1 To 3) As Integer, mCount(1 To 3) As Integer
Dim rBulk() As Integer, rHighCuts() As Integer, rLowCuts() As Integer, LowestSpec As Integer, HighestSpec As Integer, iValues() As Integer

Call FindRanges(wirename)
Call CheckWire(wirename)
grouping = FindGrouping(wirename)
SpecCuts = FindSpec(wirename)

LowestSpec = FindLowestNumber(SpecCuts)
HighestSpec = FindHighestNumber(SpecCuts)

oBulk = ListAboveNumbers(mList, HighestSpec + LowestSpec, True)
oLowCuts = ListBelowNumbers(mList, LowestSpec, False)
iValues = ListAboveNumbers(mList, LowestSpec, True)
oHighCuts = ListBelowNumbers(iValues, HighestSpec + LowestSpec, False)

check = checkarray(oBulk)
If check = True Then mCount(1) = UBound(oBulk)
check = checkarray(oHighCuts)
If check = True Then mCount(2) = UBound(oHighCuts)
check = checkarray(oLowCuts)
If check = True Then mCount(3) = UBound(oLowCuts)

rBulk = CollectRange("Bulk")
rHighCuts = CollectRange("HighCuts")
rLowCuts = CollectRange("LowCuts")

If Bulk.Rows.Count > mCount(1) Then iType(1) = 2
If Bulk.Rows.Count < mCount(1) Then iType(1) = 1
If HighCuts.Rows.Count > mCount(2) Then iType(2) = 2
If HighCuts.Rows.Count < mCount(2) Then iType(2) = 1
If LowCuts.Rows.Count > mCount(3) Then iType(3) = 2
If LowCuts.Rows.Count < mCount(3) Then iType(3) = 1

If iType(1) = 1 Then dBulk = FindDiscrepancies(oBulk, rBulk, 1)
If iType(2) = 1 Then dHighCuts = FindDiscrepancies(oHighCuts, rHighCuts, 1)
If iType(3) = 1 Then dLowCuts = FindDiscrepancies(oLowCuts, rLowCuts, 1)

If iType(1) = 2 Then dBulk = FindDiscrepancies(rBulk, oBulk, 1)
If iType(2) = 2 Then dHighCuts = FindDiscrepancies(rHighCuts, oHighCuts, 1)
If iType(3) = 2 Then dLowCuts = FindDiscrepancies(rLowCuts, oLowCuts, 1)

Call FillArray(oValues, dBulk)
Call FillArray(oValues, dHighCuts)
Call FillArray(oValues, dLowCuts)

FindAllDiscrepancies = oValues

End Function

Function MatchRange(mList, mRange, grouping)

For Each cell In mRange
 For Each arr In mList
  Match = MatchGrouping(arr, cell.Value, grouping)
  If Match = True Then
   a = a + 1
   MatchList(a) = cell.Value
   arr = ""
   Exit For
  End If
 Next arr
Next cell

SortMatchRange = MatchList
End Function


Function FindTotals(mRange)
Dim oSum As Integer, cell As Range

For Each cell In mRange
 oSum = oSum + cell.Value
Next cell

FindTotals = oSum

End Function

Function FindAllTotals()
Dim oSum As Integer, iSum(1 To 3) As Integer, a As Integer

iSum(1) = FindTotals(Bulk)
iSum(2) = FindTotals(HighCuts)
iSum(3) = FindTotals(LowCuts)

For a = 1 To 3
oSum = oSum + iSum(a)
Next a

FindAllTotals = oSum
End Function

Function CountMatchingRanges(mList, mRange, grouping)
Dim cell As Range, arr As Variant, Match As Boolean
Dim MatchList() As Integer, a As Integer, b As Integer

For Each arr In mList
 For Each cell In mRange
  Match = MatchGrouping(arr, cell.Value, grouping)
  If Match = True Then
   a = a + 1
  End If
 Next cell
 b = b + 1
 ReDim Preserve MatchList(1 To b)
 MatchList(b) = a
Next arr

CountMatchingRanges = MatchList
End Function

Sub RemoveFromRange(mRange, mList)
Dim cell As Range, arr As Variant, check As Boolean
Dim LastRow As Integer, NewLastRow As Integer

check = checkarray(mList)
If check = False Then Exit Sub

For Each arr In mList
check = mCheckRange(mRange)
If check = False Then Exit Sub
 For Each cell In mRange
  If cell.Value = arr Then
   cell.Delete (xlShiftUp)
   Exit For
  End If
 Next cell
Next arr

End Sub

Sub CheckWire(wirename)
Dim SpecCuts() As Integer, LowestSpec As Integer, check As Boolean, HighestSpec As Integer
Dim CutsBelow() As Integer, arr As Variant, cell As Range, CutsAbove() As Integer
Dim iBulk() As Integer, iHighCuts() As Integer, iLowCuts() As Integer, iValues() As Integer, iValues2() As Integer

Call FindSaveRanges(wirename)
SpecCuts = FindSpec(wirename)

LowestSpec = FindLowestNumber(SpecCuts)
HighestSpec = FindHighestNumber(SpecCuts)

iValues = ListAboveNumbers(LowCuts.Value, LowestSpec + HighestSpec, True)
Call FillArray(iBulk, iValues)
Erase iValues

iValues = ListAboveNumbers(HighCuts.Value, LowestSpec + HighestSpec, True)
Call FillArray(iBulk, iValues)
Erase iValues

iValues = ListBelowNumbers(Bulk.Value, LowestSpec + HighestSpec, False)
iValues2 = ListAboveNumbers(iValues, LowestSpec, True)
Call FillArray(iHighCuts, iValues2)
Erase iValues
Erase iValues2

iValues = ListAboveNumbers(LowCuts.Value, LowestSpec, True)
iValues2 = ListBelowNumbers(iValues, LowestSpec + HighestSpec, False)
Call FillArray(iHighCuts, iValues2)
Erase iValues
Erase iValues2

iValues = ListBelowNumbers(HighCuts.Value, LowestSpec, False)
Call FillArray(iLowCuts, iValues)
Erase iValues

iValues = ListBelowNumbers(Bulk.Value, LowestSpec, False)
Call FillArray(iLowCuts, iValues)
Erase iValues

check = checkarray(iBulk)
If check = True Then
 Call RemoveFromRange(LowCuts, iBulk)
 Call RemoveFromRange(HighCuts, iBulk)
 Call AddToRange(Bulk, iBulk)
End If

check = checkarray(iHighCuts)
If check = True Then
 Call RemoveFromRange(LowCuts, iHighCuts)
 Call RemoveFromRange(Bulk, iHighCuts)
 Call AddToRange(HighCuts, iHighCuts)
End If

check = checkarray(iLowCuts)
If check = True Then
 Call RemoveFromRange(Bulk, iLowCuts)
 Call RemoveFromRange(HighCuts, iLowCuts)
 Call AddToRange(LowCuts, iLowCuts)
End If

End Sub


Sub AddToRange(RangeName, Amounts)
Dim LastRow As Integer, NewLastRow As Integer, AddRange As Range, check As Boolean, a As Integer, cell As Range

check = checkarray(Amounts)
If check = False Then Exit Sub

Call PrintArray(Amounts)

LastRow = RangeName.Rows(RangeName.Rows.Count).Row
If RangeName.Cells(LastRow, RangeName.Columns).Value = "" Then LastRow = 2

NewLastRow = LastRow + UBound(Amounts)

Set AddRange = Range(ThisWorkbook.Worksheets(RangeName.Worksheet.Name).Cells(LastRow + 1, RangeName.Column), _
ThisWorkbook.Worksheets(RangeName.Worksheet.Name).Cells(NewLastRow, RangeName.Column))

For Each cell In AddRange
 a = a + 1
 cell.Value = Amounts(a)
Next cell

Set RangeName = Range(ThisWorkbook.Worksheets(RangeName.Worksheet.Name).Cells(RangeName.Rows(1).Row, RangeName.Column), _
ThisWorkbook.Worksheets(RangeName.Worksheet.Name).Cells(NewLastRow, RangeName.Column))

End Sub

Function CollectRange(mRange)
Dim oValues() As Integer

If mRange = "LowCuts" Then
 oValues = CollectRangeValues(LowCuts)
 CollectRange = oValues
End If
If mRange = "HighCuts" Then
 oValues = CollectRangeValues(HighCuts)
 CollectRange = oValues
End If
If mRange = "Bulk" Then
 oValues = CollectRangeValues(Bulk)
 CollectRange = oValues
End If
End Function

Function CollectRangeValues(mRange)
Dim cell As Range, oValues() As Integer, a As Integer

ReDim oValues(1 To mRange.Rows.Count)
For Each cell In mRange
 a = a + 1
 oValues(a) = cell.Value
Next cell

CollectRangeValues = oValues

End Function

Sub SortAdd(Amount, wirename)
Dim arr As Variant, oHighCuts() As Integer, oLowCuts() As Integer, oBulk() As Integer, iValues() As Integer
Dim SpecCuts() As Integer, LowestSpec As Integer, HighestSpec As Integer

Call FindRanges(wirename)
Call CheckWire(wirename)

SpecCuts = FindSpec(wirename)
LowestSpec = FindLowestNumber(SpecCuts)
HighestSpec = FindHighestNumber(SpecCuts)

oBulk = ListAboveNumbers(Amount, HighestSpec + LowestSpec, True)
oLowCuts = ListBelowNumbers(Amount, LowestSpec, False)
iValues = ListAboveNumbers(Amount, LowestSpec, True)
oHighCuts = ListBelowNumbers(iValues, HighestSpec + LowestSpec, False)

Call AddToRange(Bulk, oBulk)
Call AddToRange(LowCuts, oLowCuts)
Call AddToRange(HighCuts, oHighCuts)

End Sub

Sub SortRemove(Amount, wirename)
Dim arr As Variant, oHighCuts() As Integer, oLowCuts() As Integer, oBulk() As Integer, iValues() As Integer
Dim SpecCuts() As Integer, LowestSpec As Integer, HighestSpec As Integer

Call FindRanges(wirename)
Call CheckWire(wirename)

SpecCuts = FindSpec(wirename)
LowestSpec = FindLowestNumber(SpecCuts)
HighestSpec = FindHighestNumber(SpecCuts)

oBulk = ListAboveNumbers(Amount, HighestSpec + LowestSpec, True)
oLowCuts = ListBelowNumbers(Amount, LowestSpec, False)
iValues = ListAboveNumbers(Amount, LowestSpec, True)
oHighCuts = ListBelowNumbers(iValues, HighestSpec + LowestSpec, False)

Call RemoveFromRange(Bulk, oBulk)
Call RemoveFromRange(LowCuts, oLowCuts)
Call RemoveFromRange(HighCuts, oHighCuts)

End Sub

Sub iRemoveWireType(wirename)

Call SetWorksheets(hc, "HIGH CUT")
Call SetWorksheets(lc, "LOW CUT")
Call SetWorksheets(bk, "BULK")
Call SetRange(wirename, Bulk, bk, True)
Call SetRange(wirename, HighCuts, hc, True)
Call SetRange(wirename, LowCuts, lc, True)
Bulk.Delete (xlShiftToLeft)
HighCuts.Delete (xlShiftToLeft)
LowCuts.Delete (xlShiftToLeft)

End Sub

Function CheckLow(wirename)
Dim iValues() As Integer, iOrders As Integer, iAverage As Integer

Call FindRanges(wirename)
iOrders = FindOrders(wirename)
iValues = FindAllNumberOfIncrements(wirename)
iAverage = FindAverage(iValues)

If iAverage <= iOrders Then CheckLow = True
If iAverage >= iOrders Then CheckLow = False

End Function

Function iCollectLowWire()
Dim WireType() As Variant, a As Integer, oValues() As Variant, Low As Boolean, b As Integer
WireType = sFindAllWireTypes

For a = 1 To UBound(WireType)
Low = False
Low = CheckLow(WireType(a))
If Low = True Then
 b = b + 1
 ReDim Preserve oValues(1 To b)
 oValues(b) = WireType(a)
End If
End Function

Function FindNumberOfIncrements(mValue)
Dim cell As Range, oValue As Integer

For Each cell In Bulk
 If cell.Value > mValue Then
  oValue = oValue + cell.Value \ mValue
 End If
Next cell
For Each cell In HighCuts
 If cell.Value > mValue Then
  oValue = oValue + cell.Value \ mValue
 End If
Next cell

FindNumberOfIncrements = oValue

End Function

Function FindAllNumberOfIncrements(wirename)
Dim a As Integer, iValues() As Integer, oValues() As Integer

iValues = FindSpec(wirename)
ReDim oValues(1 To UBound(iValues))
iValues = SortLowToHigh(iValues)

For a = 1 To UBound(iValues)
 oValues(a) = FindNumberOfIncrements(iValues(a))
Next a
 
FindAllNumberOfIncrements = oValues

End Function

Sub VerifyCount(mList, wirename)
Dim iValues() As Integer, CheckA As Boolean, answer As String, a As Integer, sum As Integer
Dim iRangeValues() As Integer, b As Integer

ReDim iRangeValues(1 To View.vcInvList.ListCount)
For a = 1 To View.vcInvList.ListCount
 iRangeValues(a) = View.vcInvList.list(b)
 b = b + 1
Next a

iValues = FindAllDiscrepancies(mList, wirename)
CheckA = checkarray(iValues)
If CheckA = True Then
 For a = 1 To UBound(iValues)
  View.vcDiffList.object.AddItem iValues(a)
  sum = sum + iValues(a)
 Next a
End If
View.vcDiffTotal.Caption = sum
If CheckA = False Then MsgBox ("No discrepancies found in count")



End Sub

Function invFindRangeValues(wirename, mString)
Dim a As Integer, b As Integer, iValues() As Integer, mRange As Range

Call FindRanges(wirename)
Call FindRangeName(mString, wirename, mRange)

For a = 1 To mRange.Rows.Count
 b = b + 1
 ReDim Preserve iValues(1 To b)
 iValues(b) = mRange.Rows(a)
Next a

invFindRangeValues = iValues

End Function

Function invFindAllRangeValues(wirename)
Dim oBulk() As Integer, oHighCuts() As Integer, oLowCuts() As Integer, a As Integer, b As Integer, CheckA As Boolean
Dim oValues() As Integer

Call FindRanges(wirename)
oBulk = invFindRangeValues(wirename, "Bulk")
oHighCuts = invFindRangeValues(wirename, "HighCuts")
oLowCuts = invFindRangeValues(wirename, "LowCuts")

CheckA = checkarray(oBulk)
If CheckA = True Then
For a = 1 To UBound(oBulk)
 b = b + 1
 ReDim Preserve oValues(1 To b)
 oValues(b) = oBulk(a)
Next a
End If

CheckA = checkarray(oHighCuts)
If CheckA = True Then
For a = 1 To UBound(oHighCuts)
 b = b + 1
 ReDim Preserve oValues(1 To b)
 oValues(b) = oHighCuts(a)
Next a
End If

CheckA = checkarray(oLowCuts)
If CheckA = True Then
For a = 1 To UBound(oLowCuts)
 b = b + 1
 ReDim Preserve oValues(1 To b)
 oValues(b) = oLowCuts(a)
Next a
End If
invFindAllRangeValues = oValues

End Function

Sub FindRangeName(mString, wirename, mRange)
If mString = "Bulk" Then Call SetRange(wirename, mRange, bk, False)
If mString = "HighCuts" Then Call SetRange(wirename, mRange, hc, False)
If mString = "LowCuts" Then Call SetRange(wirename, mRange, lc, False)
End Sub
