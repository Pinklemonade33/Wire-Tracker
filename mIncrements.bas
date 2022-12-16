Attribute VB_Name = "mIncrements"
Option Explicit


Sub FindIncrements(wirename, IncAmount, grouping, IncQTY)
Dim LowCuts() As Integer, HighCuts() As Integer, Bulk() As Integer, arr As Variant, Sums() As Integer
Dim Check0 As Boolean, a As Integer, HighestSum As Integer

Call FindRanges(wirename)
LowCuts = FindAboveAndBelow(IncAmount, grouping, "LowCuts", wirename)
HighCuts = FindAboveAndBelow(IncAmount, grouping, "HighCuts", wirename)
Bulk = FindAboveAndBelow(IncAmount, grouping, "Bulk", wirename)

LowCuts = SortLowToHigh(LowCuts)
HighCuts = SortLowToHigh(HighCuts)
Bulk = SortLowToHigh(Bulk)

ReDim Preserve LowCuts(1 To IncQTY)
ReDim Preserve HighCuts(1 To IncQTY)
ReDim Preserve Bulk(1 To IncQTY)

Check0 = CheckForKeyword(0, LowCuts)
If Check0 = False Then
 a = a + 1
 ReDim Preserve Sums(1 To a)
 Sums(a) = SumArray(LowCuts)
End If
Check0 = CheckForKeyword(0, HighCuts)
If Check0 = False Then
 a = a + 1
 ReDim Preserve Sums(1 To a)
 Sums(a) = SumArray(HighCuts)
End If
Check0 = CheckForKeyword(0, Bulk)
If Check0 = False Then
 a = a + 1
 ReDim Preserve Sums(1 To a)
 Sums(a) = SumArray(Bulk)
End If

Check0 = checkarray(Sums)
If Check0 = True Then HighestSum = FindLowestNumber(Sums)

End Sub

Function FindAboveAndBelow(IncAmount, grouping, mRange, wirename)
Dim InputNumbers() As Integer, arr As Variant, OutputNumbers() As Integer

OutputNumbers = CollectRange(mRange)

InputNumbers = OutputNumbers
Erase OutputNumbers
OutputNumbers = ListAboveNumbers(InputNumbers, IncAmount, True)

InputNumbers = OutputNumbers
Erase OutputNumbers
OutputNumbers = ListBelowNumbers(InputNumbers, IncAmount + grouping, True)

FindAboveAndBelow = OutputNumbers

End Function

Sub test66()
Call FindIncrements("6L2GRY", 180, 30, 2)

End Sub
