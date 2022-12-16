Attribute VB_Name = "mArrayTools"
Option Explicit

Function aFindAll(mList, KeyWord, mType)
Dim oValues() As Integer, oRanges() As Integer
Dim a As Integer

For a = 1 To UBound(mList)
 If mList(a) = KeyWord Then
  b = b + 1
  ReDim Preserve Output(1 To b)
  Output(b) = mList(a + 1)
 End If
Next a

If mType = "Values" Then aFindAll = oValues
If mType = "Ranges" Then aFindAll = oRanges

End Function

Function aFindRangesTwoWay(mStart, mList, KeyStart, KeyEnd)
Dim a As Integer, Output As Integer

For a = mStart To 1 Step -1
 If mList(a) = KeyStart Then Output(1) = a
Next a

For a = mStart To UBound(mList)
 If mList(a) = KeyEnd Then Output(2) = a
Next a

aFindRangesTwoWay = Output(1)
aFindRangesTwoWay = Output(2)

End Function


Function aFindRanges(mStart, mList, mEnd, KeyStart, KeyEnd)
Dim a As Integer, Output(1 To 2) As Variant

For a = mStart To mEnd
 If mList(a) = KeyStart Then Output(1) = a
 If mList(a) = KeyEnd Then Output(2) = a
Next a

aFindRanges = Output

End Function

Function aCollectValues(mStart, mEnd, mList)
Dim a As Integer, b As Integer, oValues() As Variant

For a = mStart + 1 To mEnd + -1
 If Not a = mStart Or Not a = mEnd Then
  b = b + 1
  ReDim Preserve oValues(1 To b)
  oValues(b) = mList(a)
 End If
Next a

aCollectValues = oValues

End Function


Sub test33()
Dim a() As Variant, arr As Variant, b As Integer, c() As Variant, d() As Variant, aa As Integer

ReDim a(1 To 10)
For aa = 1 To 10
 b = b + 1
 a(aa) = b
Next aa
For Each arr In a
Debug.Print arr
Next arr
a(3) = "start"
a(9) = "end"

c = aFindRanges(1, a, 10, "start", "end")
For Each arr In a
Debug.Print arr
Next arr
d = aCollectValues(c(1), c(2), a)

End Sub
