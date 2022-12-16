Attribute VB_Name = "mMisc"
Option Explicit

Function TakeFromArray(tArray, FirstKeyWord, LastKeyWord)
Dim a As Integer, b As Integer, OutputArr() As Variant

For a = 1 To UBound(tArray)
 If tArray(a) = FirstKeyWord Then
  Do Until tArray(a) = LastKeyWord
   a = a + 1
   b = b + 1
   ReDim Preserve OutputArr(1 To c)
   OutputArr(c) = tArray(a)
  Loop
  Exit For
 End If
Next a

TakeFromArray = tArray

End Function

Function MatchGrouping(Primary, Secondary, grouping)
    If Primary >= Secondary And Primary <= Secondary + grouping Then
      MatchGrouping = True
      Else: MatchGrouping = False
    End If
End Function

Function FindLowestNumber(mList)
Dim arr As Variant, LowestNumber As Integer, First As Boolean, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then Exit Function

First = True
For Each arr In mList
 If arr < LowestNumber Or First = True Then LowestNumber = arr
 First = False
Next arr

FindLowestNumber = LowestNumber
End Function

Function FindHighestNumber(mList)
Dim arr As Variant, HighestNumber As Integer, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then Exit Function

For Each arr In mList
 If arr > HighestNumber Then HighestNumber = arr
Next arr

FindHighestNumber = HighestNumber
End Function

Function ListBelowNumbers(mList, Number, equal)
Dim arr As Variant, Below() As Integer, a As Integer, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then
 ListBelowNumbers = Below
 Exit Function
End If

If equal = True Then Number = Number + 1

For Each arr In mList
 If arr < Number Then
  a = a + 1
  ReDim Preserve Below(1 To a)
  Below(a) = arr
 End If
Next arr

ListBelowNumbers = Below

End Function

Function ListAboveNumbers(mList, Number, equal)
Dim arr As Variant, Above() As Integer, a As Integer, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then
 ListAboveNumbers = Above
 Exit Function
End If

If equal = True Then Number = Number + -1

For Each arr In mList
 If arr > Number Then
  a = a + 1
  ReDim Preserve Above(1 To a)
  Above(a) = arr
 End If
Next arr


ListAboveNumbers = Above
End Function

Function SortHighToLow(mList)
Dim a As Integer, Low As Integer, High As Integer, good As Boolean, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then
 SortHighToLow = mList
 Exit Function
End If

good = False
Do Until good = True
good = True
For a = 1 To UBound(mList)
 If Not a + 1 > UBound(mList) Then
  If mList(a + 1) > mList(a) Then
    High = mList(a + 1)
    Low = mList(a)
    mList(a + 1) = Low
    mList(a) = High
    good = False
  End If
 End If
Next a
Loop

SortHighToLow = mList
 
End Function

Function SortLowToHigh(mList)
Dim a As Integer, Low As Integer, High As Integer, good As Boolean, checkb As Boolean
checkb = checkarray(mList)
If checkb = False Then
 SortLowToHigh = mList
 Exit Function
End If

good = False
Do Until good = True
good = True
For a = 1 To UBound(mList)
 If Not a + 1 > UBound(mList) Then
  If mList(a + 1) < mList(a) Then
    Low = mList(a + 1)
    High = mList(a)
    mList(a + 1) = High
    mList(a) = Low
    good = False
  End If
 End If
Next a
Loop

SortLowToHigh = mList
 
End Function

Function TransferArray(Primary, Secondary)
Dim a As Integer
ReDim Primary(1 To UBound(Secondary))
For a = 1 To UBound(Primary)
 Primary(a) = Secondary(a)
Next a

TransferArray = Primary

End Function

Function SumArray(mList)
Dim arr As Variant, sum As Integer

For Each arr In mList
 sum = sum + arr
Next arr

SumArray = sum

End Function

Function CheckForKeyword(KeyWord, mList)
Dim arr As Variant

For Each arr In mList
 If arr = KeyWord Then
  CheckForKeyword = True
  Exit Function
 End If
Next arr

CheckForKeyword = False
End Function

Function FindLastRow(mWS, KeyEnd, mCol, mRow)

Do Until mWS.Cells(mRow + 1, mCol).Value = KeyEnd
 mRow = mRow + 1
Loop
FindLastRow = mRow

End Function

Function FindLastRowForRange(mRange, mRow)
FindLastRowForRange = mRow + mRange.Rows.Count + -1
End Function

Sub BordersColorsFontsValues(mRange, BorderType, Red, Green, Blue, Position, tSize, tColor, iValue, rHeight)
mRange.Color = RGB(Red, Green, Blue)
mRange.Value = iValue
Row.Height = rHeight

If BorderType = "Standard" Then
  mRange.Borders.LineStyle = LineStyle.XlLineStyle.xlContinuous
End If

If BorderType = "Header" Then
  mRange.Borders.LineStyle = LineStyle.XlLineStyle.xlContinuous
  mRange.Borders(xlBottom).Weight = xlThick
  mRange.Borders(xlTop).Weight = xlThick
End If

If BorderType = "Side" Then
 mRange.Border.LineStyle = LineStyle.XlLineStyle.xlContinuous
 mRange.Borders(xlRight).Weight = xlThick
 If Position = "Top" Then mRange.Borders(xlTop).Weight = xlThick
 If Position = "Bottom" Then mRange.Borders(xlBottom).Weight = xlThick
End If

If BorderType = "wType" Then
 mRange.Border.LineStyle = LineStyle.XlLineStyle.xlContinuous
 mRange.Borders(xlBottom).Weight = xlThick
End If
 
If tColor = "White" Then mRange.Font.ColorIndex = -4142
If tColor = "Black" Then mRange.Font.Color = 1

Font.Size = tSize

End Sub

Function ImportRange(mRange)
 ImportRange = pExportRange(mRange)
End Function

Function FindAverage(mList)
Dim iValue As Integer

iValue = SumArray(mList)
FindAverage = iValue \ UBound(mList)

End Function

Function FindDiscrepancies(Primary, Secondary, oType)
Dim Status() As String, a As Integer, b As Integer, check As Boolean, Status2() As String, oValues()

ReDim Status(1 To UBound(Primary))
check = checkarray(Secondary)
If check = True Then ReDim Status2(1 To UBound(Secondary))

If check = True Then
For a = 1 To UBound(Primary)
 For b = 1 To UBound(Secondary)
  If Primary(a) = Secondary(b) And Not Status2(b) = "Match" Then
   Status(a) = "Match"
   Status2(b) = "Match"
   Exit For
  End If
 Next b
Next a
End If

b = 0
For a = 1 To UBound(Status)
 If oType = 1 Then
  If Status(a) = "" Then
   b = b + 1
   ReDim Preserve oValues(1 To b)
   oValues(b) = Primary(a)
  End If
 End If
 If oType = 2 Then
  If Status(a) = "Match" Then
   b = b + 1
   ReDim Preserve oValues(1 To b)
   oValues(b) = Primary(a)
  End If
 End If
Next a

FindDiscrepancies = oValues

End Function

Sub FillArray(Primary, Secondary)
Dim a As Integer, b As Integer, check As Boolean

check = checkarray(Secondary)
If check = False Then Exit Sub

check = False
check = checkarray(Primary)
If check = True Then b = UBound(Primary)

For a = 1 To UBound(Secondary)
 b = b + 1
 ReDim Preserve Primary(1 To b)
 Primary(b) = Secondary(a)
Next a

End Sub

Sub PrintArray(mList)
Dim a As Integer

For a = 1 To UBound(mList)
 Debug.Print mList(a)
Next a

End Sub

Function checkarray(a) As Boolean
    On Error Resume Next
    checkarray = Not LBound(a) > UBound(a)
End Function

Function mCheckRange(mRange) As Boolean
Dim test As Range
 On Error Resume Next
 If mRange.Rows Is Nothing Then
  mCheckRange = False
  Else: mCheckRange = True
 End If
End Function

Sub FillObject(mObject, mList)
Dim a As Integer

For a = 1 To UBound(mList)
 mObject.object.AddItem mList(a)
Next a

End Sub

Function SortLongIntoArray(mLong)
Dim a As Integer, LastRange As Integer, pStart As Integer, pEnd As Integer, oValues() As Variant
Dim b As Integer, c As Integer

LastRange = Len(mLong)

For a = 1 To LastRange
 If Mid(mLong, a, 1) > " " Then
  pStart = a
  b = 1
  Do Until Mid(mLong, a + 1, 1) = " " Or Mid(mLong, a + 1, 1) = ""
   a = a + 1
   b = b + 1
  Loop
  pEnd = b
  c = c + 1
  ReDim Preserve oValues(1 To c)
  oValues(c) = Mid(mLong, pStart, pEnd)
 End If
Next a

SortLongIntoArray = oValues

End Function

Function FindFullWord(mValue, mList)
Dim a As Integer, b As Integer, c As Integer, oValues() As Variant, LowerCase As Variant

For a = 1 To UBound(mList)
LowerCase = LCase(mList(a))
 For b = 1 To Len(mList(a))
    If Mid(mList(a), b, b + Len(mValue) + -1) = mValue _
    Or Mid(LowerCase, b, b + Len(mValue) + -1) = mValue Then
      c = c + 1
      ReDim Preserve oValues(1 To c)
      oValues(c) = mList(a)
    End If
 Next b
Next a

FindFullWord = oValues
  
End Function

Function CheckValidation(mCell)
On Error Resume Next
If IsEmpty(mCell.Validation.Type) Then
 CheckValidation = False
 Else: CheckValidation = True
End If
End Function

Function FindDuplicates(mList)
Dim a As Integer, oValues(), b As Integer, c As Integer, d As Integer, check As Boolean, First As Boolean, Dcount As Integer
First = True

check = checkarray(mList)
If check = False Then
 FindDuplicates = oValues
 Exit Function
End If

For a = 1 To UBound(mList)
 For b = 1 To UBound(mList)
  If mList(a) = mList(b) And First = True Then
    ReDim oValues(1 To 1)
    oValues(1) = mList(a)
    d = 1
    First = False
    Exit For
  End If
   Dcount = 0
   For c = 1 To UBound(oValues)
    If mList(a) = oValues(c) Then Dcount = Dcount + 1
   Next c
   If Dcount > 1 And mList(a) = mList(b) Then
    d = d + 1
    ReDim Preserve oValues(1 To d)
    oValues(d) = mList(a)
   End If
 Next b
Next a
    
FindDuplicates = oValues

End Function

