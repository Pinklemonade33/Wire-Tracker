Attribute VB_Name = "mPrevCalc"
Option Explicit
Dim StoredP As Range, ActiveP As Range, EmptyP As Range
Dim WireType As Range, FromPreCuts As Range, FromSpool As Range, DateRange As Range
Dim AdditionalCuts As Range, Pending As Range, FromLap As Range, WireToSites As Range, SiteID As Range
Dim LapRequestedCuts As Range, LapAddCuts As Range, LapBeforeSpool As Range, LapAfterSpool As Range, SiteWire As Range

Function pCollectLap(mString, Lap)
mRange As Range
mRange = FindRangeName(mString)
pCollectLap = CollectLap(mRange, Lap)
End Function

Function pCollectTotals(mString, wirename)
Dim mRange As Range
mRange = FindRangeName(mString)
Call FindProjectionRange(True, 0)
pCollectTotals = Collect(wirename, ActiveP, mRange)
End Function

Sub StoreProjection()
Dim a As Integer, FirstRow As Integer, LastRow As Integer, LastCol As Integer
Dim pFirstRow As Integer, pLastRow As Integer, ws As Worksheet
Set ws = ThisWorkbook.Worksheets("StoredProjections")

Call FindProjectionRange(False, 0)
FirstRow = FindNewProjection
pFirstRow = FirstRow + 1
pLastRow = FindLastRowForRange(ActiveP, pFirstRow)
LastRow = pLastRow + 1
LastCol = ActiveP.Columns.Count
Call pSetRange(pFirstRow, pLastRow, 1, LastCol, ws, EmptyP)

ws.Cells(FirstRow, 1).Value = "pNumber"
a = FindLastProjectionNumber
ws.Cells(FirstRow, 1).Value = a + 1
ws.Cells(pFirstRow, LastCol).Value = "LastCol"
ws.Cells(pFirstRow, LastCol).Font.ColorIndex = -4142
EmptyP = ActiveP
ws.Cells(LastRow, 1).Value = "LastRow"
ws.Cells(LastRow, 2).Value = Now
ws.Cells(LastRow + 1, 1).Value = "NewProjection"

End Sub

Function CollectLap(mRange, Lap)
Dim cell As Range, oValues() As Integer, a As Integer
Dim StartKey As Variant, EndKey As Variant

Call FindProjectionRange(True, 0)
StartKey = FindStartKey(mRange)
EndKey = FindEndKey(mRange)
Call FindLap(Lap, ActiveP)
Call DetailRange(ActiveP, StartKey, EndKey, mRange, False)

For Each cell In mRange
 a = a + 1
 ReDim Preserve oValues(1 To a)
 oValues(a) = cell.Value
Next cell

CollectLap = oValues

End Function

Function CollectAllWireType(wirename, mdrange)
Dim LastRange As Integer, a As Integer, iValues() As Integer
Dim b As Integer, c As Integer

LastRange = FindLastProjectionNumber
StartKey = FindStartKey(mType)
EndKey = FindEndKey(mType)

For a = 1 To LastRange
 Call FindProjectionRange(False, a)
 iValues = pCollect(wirename, StoredP, mdrange)
  For b = 1 To UBound(iValues)
   c = c + 1
   ReDim Preserve oValues(1 To c)
   oValues(c) = iValues(b)
  Next b
Next a

CollectAllWireType = oValues

End Function

Sub FindLap(Lap, mRange)
Dim cell As Range, FirstRow As Integer, LastRow As Integer

For Each cell In Range(mRange.Cells(1, 1), mRange.Cells(mRange.Rows.Count, 1))
 If cell.Value = "LAP " & Lap Then FirstRow = cell.Row
 If cell.Value = "LAP " & Lap + 1 Or cell.Value = "TOTALS" Then LastRow = cell.Row + -1
Next cell

Call pSetRange(FirstRow, LastRow, 1, mRange.Columns, mRange.Worksheet, FromLap)
End Sub

Function Collect(wirename, mRange, mdrange)
Dim cell As Range, oValues() As Integer, a As Integer
Dim StartKey As Variant, EndKey As Variant

StartKey = FindStartKey(mdrange)
EndKey = FindEndKey(mdrange)

Call DetailRange(mRange, StartKey, StartKey, mdrange, wirename)
For Each cell In mdrange
 a = a + 1
 ReDim Preserve oValues(1 To a)
 oValues(a) = cell.Value
Next cell

pCollect = oValues

End Function

Function FindWireTypeColumn(wirename)
Dim cell As Range
Call DetailRange(mRange, "Wire Type", "From Pre-Cuts", WireType, False)

For Each cell In WireType
 If cell.Value = wirename Then
  CheckForWireType = cell.Column
  Exit Function
 End If
Next cell

End Function

Sub Recount()
Dim a As Integer, ws As Worksheet
Set ws = ThisWorkbook.Worksheets("StoredProjections")

Do Until Cells(a, 1).Value = "NewProjection"
a = a + 1
 If Cells(a, 1).Value = "pNumber" Then
  b = b + 1
  Cells(a, 2).Value = b
 End If
Loop

End Sub

Function FindNewProjection()
Dim a As Integer, ws As Worksheet
Set ws = ThisWorkbook.Worksheets("StoredProjections")

Do
a = a + 1
If ws.Cells(a, 1).Value = "NewProjection" Then
 FindNewProjection = a
 Exit Function
End If
Loop

End Function

Sub FindProjectionRange(Active, mNumber)
Dim a As Integer, ws As Worksheet, FirstRow As Integer, LastRow As Integer, LastCol As Integer

If Active = True Then Set ws = ThisWorkbook.Worksheets("Projection")
If Active = False Then Set ws = ThisWorkbook.Worksheets("StoredProjections")

If Active = False Then
    Do
      a = a + 1
      If ws.Cells(a, 1).Value = "pNumber" _
      And ws.Cells(a, 2).Value = mNumber Then Exit Do
    Loop
End If

Do Until ws.Cells(a + 1, b + 1).Value = "LastCol"
 b = b + 1
Loop
 LastCol = b
 
FirstRow = a + 1
Do
 a = a + 1
 If ws.Cells(a + 1, 1).Value = "LastRow" Then LastRow = a
Loop

If Active = False Then Call pSetRange(FirstRow, LastRow, 1, LastCol, ws, StoredP)
If Active = True Then Call pSetRange(FirstRow, LastRow, 1, LastCol, ws, ActiveP)

End Sub

Function FindLastProjectionNumber()
Dim a As Integer, ws As Worksheet
Set ws = ThisWorkbook.Worksheets("StoredProjections")

a = FindNewProjection

Do
a = a + -1
If ws.Cells(a, 1).Value = "pNumber" Then
 FindLastProjection = ws.Cells(a, 2).Value
 Exit Function
End If
Loop

End Function


Sub DetailRange(mRange, StartKey, EndKey, nRange, wirename)
Dim cell As Range, LastRow As Integer, ws As String, FirstCol As Integer, LastCol As Integer, a As Integer

a = 1

If Not wirename = False Then
 FirstCol = FindWireTypeColumn(wirename)
 LastCol = FirstCol
End If

If wirename = False Then
 FirstCol = 2
 LastCol = mRange.Columns.Count
End If

If wirename = "SiteID" Then
 FirstCol = 2
 LastCol = 2
 a = 2
End If

If wirename = "Date" Then
 FirstCol = 2
 LastCol = 2
 a = 1
End If

For Each cell In Range(mRange.Cells(1, a), mRange.Cells(mRange.Rows.Count, a))
 If cell.Value = StartKey Then
  LastRow = FindLastRow(mRange.Worksheet, EndKey, a, cell.Row)
  Call pSetRange(mRange.Row, LastRow, FirstCol, LastCol, mRange.Worksheet.Name, nRange)
  Exit Sub
 End If
Next cell

End Sub

Sub SetRange(FirstRow, LastRow, FirstCol, LastCol, Worksheet, mRange)
Set mRange = Range(ThisWorkbook.Worksheets(Worksheet).Cells(FirstRow, FirstCol), ThisWorkbook.Worksheets(Worksheet).Cells(LastRow, LastCol))
End Sub

Function FindCurrentSites(wirename, mType)
Dim LastRange As Integer, a As Integer, pNumbers() As Integer, Sites() As Variant
Dim cell As Range, b As Integer, c As Integer, good As Boolean, First As Boolean

LastRange = FindLastProjectionNumber
First = True

For a = 1 To LastRange
 Call FindProjectionRange(False, a)
 Call DetailRange(StoredP, "Wire to Sites", "LastRow", WireToSites, wirename)
 Call DetailRange(WireToSites, "SiteID", "", SiteID, "SiteID")
  For Each cell In SiteID
   good = True
   If First = False Then
    For b = 1 To UBound(Sites)
     If Sites(b) = cell.Value Then
      good = False
      pNumbers(b) = StoredP.Cells(1, 1).Value
      Exit For
    Next b
   End If
   If good = True Or First = True Then
    First = False
    c = c + 1
    ReDim Preserve Sites(1 To c)
    ReDim Preserve pNumbers(1 To c)
    Sites(c) = cell.Value
    pNumbers(c) = StoredP.Cells(1, 1).Value
   End If
  Next cell
Next a

If mType = "Sites" Then FindCurrentSites = Sites
If mType = "pNumbers" Then FindCurrentSites = pNumbers

End Function

Function pCollectFromSites(wirename, mSites, mType, MaxL, MinL, MaxD, MinD, mType2)
Dim a As Integer, iSites() As Variant, pNumbers() As Integer, iNumbers() As Integer, oDates() As Date, Sites() As Integer
Dim Colors() As Integer, oValues() As Integer, oSites() As Variant, b As Integer

iSites = FindCurrentSites(wirename, "Sites")
iNumbers = FindCurrentSites(wirename, "pNumbers")
Colors = FindType(mType)

If mSites = False Then
 pNumbers = iNumbers
 Sites = iSites
End If

If Not mSites = False Then
 pNumbers = SortpNumbers(iNumbers, mSites, iSites)
 Sites = mSites
End If

For a = 1 To UBound(pNumbers)
 Call FindProjectionRange(False, pNumber)
 Call DetailRange(StoredP, "Wire to Sites", "", WireToSites, False)
 Call DetailRange(WireToSites, "SiteID", "", SiteID, False)
 Call DetailRange(SiteID, mSites(a), mSites(a), SiteWire, wirename)
 Call DetailRange(StoredP, "LastRange", "LastRange", DateRange, "Date")
 If SiteWire.Interior.Color = RGB(Colors(1), Colors(2), Colors(3)) Or cell.Interior.Color = RGB(Colors(4), Colors(5), Colors(6)) Then
  If SiteWire.Value >= MinL And SiteWire.Value <= MaxL Or MaxL = False And SiteWire.Value = MinL Then
   b = b + 1
   ReDim Preserve oValues(1 To b)
   ReDim Preserve oSites(1 To b)
   ReDim Preserve oDates(1 To b)
   oValues(b) = SiteWire.Value
   If StoredP.Value >= MinD And StoredP.Value <= MaxD Or MaxD = False And DateRange.Value = MinD Then oDates(b) = StoredP.Cells(StoredP.Rows.Count, 2).Value
   oSites(b) = Sites(a)
 End If
Next a

If mType2 = "Values" Then pCollectFromSites = oValues
If mType2 = "Dates" Then pCollectFromSites = oDates
If mType2 = "Sites" Then pCollectFrom Sites = oSites

End Function

Function FindType(mType)
Dim oValues(1 To 6) As Integer
If mType = "Pending" Then
 oValues(1) = 255
 oValues(2) = 80
 oValues(3) = 80
End If
If mytype = "Picked" Then
 oValues(1) = 146
 oValues(2) = 208
 oValues(3) = 80
 oValues(4) = 0
 oValues(5) = 176
 oValues(6) = 80
End If

FindType = oValues
End Function

Function SortpNumbers(iNumbers, mSites, iSites)
Dim a As Integer, b As Integer, pNumbers() As Integer

ReDim pNumbers(1 To UBound(iSites))
For a = 1 To UBound(iSites)
 For b = 1 To UBound(mSites)
  If mSites(b) = iSites(a) Then
   pNumbers(b) = iNumbers(a)
   Exit For
  End If
 Next b
Next a
 
SortpNumbers = pNumbers

End Function

Function FindRangeName(mString)
If mString = "WireToSites" Then FindRangeName = WireToSites
If mString = "WireType" Then FindRangeName = WireType
If mString = "PreCuts" Then FindRangeName = PreCuts
If mString = "FromSpool" Then FindRangeName = FromSpool
If mString = "Pending" Then FindRangeName = Pending
If mString = "AdditionalCuts" Then FindRangeName = AdditionalCuts
If mString = "FromLap" Then FindRangeName = FromLap
If mString = "LapRequestedCuts" Then FindRangeName = LapRequestedCuts
If mString = "LapAddCuts" Then FindRangeName = LapAddCuts
If mString = "LapBeforeSpool" Then FindRangeName = LapBeforeSpool
If mString = "AfterSpool" Then FindRangeName = LapAfterSpool
If mString = "FromLap" Then FindRangeName = FromLap
End Function


Function FindStartKey(mRange)
If mRange = WireToSites Then FindStartKey = "Wire to Sites"
If mRange = WireType Then FindStartKey = "Wire Type"
If mRange = PreCuts Then FindStartKey = "From Pre-Cuts"
If mRange = FromSpool Then FindStartKey = "From Spool"
If mRange = Pending Then FindStartKey = "Pending"
If mRange = AdditionalCuts Then FindStartKey = "Additional Cuts"
If mRange = LapRequestedCuts Then FindStartKey = "Requested Cuts"
If mRange = LapAddCuts Then FindStartKey = "Additional Cuts"
If mRange = LapBeforeSpool Then FindStartKey = "Initial Spool"
If mRange = LapAfterSpool Then FindStartKey = "Resulting Spool"
End Function

Function FindEndKey(mRange)
If mRange = WireToSites Then FindEndKey = "LastRow"
If mRange = WireType Then FindEndKey = "From Pre-Cuts"
If mRange = PreCuts Then FindEndKey = "From Spool"
If mRange = FromSpool Then FindEndKey = "Pending"
If mRange = Pending Then FindEndKey = "Additional Cuts"
If mRange = AdditionalCuts Then FindEndKey = "Sites"
If mRange = LapRequestedCuts Then FindEndKey = "Additional Cuts"
If mRange = LapAddCuts Then FindEndKey = "LAP"
If mRange = LapBeforeSpool Then FindEndKey = "Sum Of Cuts"
If mRange = LapAfterSpool Then FindEndKey = "Requested Cuts"
End Function

Function pExportRange(mRange)
Dim iRange As Range, StartKey As Variant, EndKey As Variant
Dim oValues(1 To 4) As Integer

iRange = FindRangeName(mRange)
StartKey = FindStartKey(iRange)
EndKey = FindEndKey(iRange)

Call DetailRange(ActiveP, StartKey, EndKey, iRange, False)

oValues(1) = iRange.Row
oValues(2) = iRange.Rows(iRange.Rows.Count).Row
oValues(1) = iRange.Column
oValues(2) = iRange.Columns(iRange.Columns.Count).Column

pExportRange = oValues

End Function

