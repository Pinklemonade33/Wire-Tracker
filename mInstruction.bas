Attribute VB_Name = "mInstruction"
Option Explicit
Dim ws1 As Worksheet, ws2 As Worksheet
Dim PrimaryCuts As Range, BeforeSpool As Range, AfterSpool As Range, SecondaryCuts As Range
Dim WireType As Range, LapRange As Range, SiteRange As Range
Dim SideWireType As Range, SideBefore As Range, SideAfter As Range, SidePrimCuts As Range, SideSecCuts As Range
Dim DirWireType As Range, DirRemove As Range, DirAdd As Range, DirLap As Range
Dim DirSideType As Range, DirSideRemove As Range, DirSideAdd As Range


Sub SetRange(mRange, FirstRow, LastRow, FirstCol, LastCol, ws)
Set mRange = Range(ws.Cells(FirstRow, LastRow), ws.Cells(LastRow, LastCol))
End Sub

Function FindLapRanges(RequestedCuts, AddCuts)
Dim a As Integer, ws As Worksheet, oValues(1 To 6) As Integer
Dim LapRow As Integer, TypeRow As Integer, BeforeRow As Integer, ReqRow As Integer, AddRow As Integer, AfterRow As Integer
Set ws = ThisWorkbook.Worksheets("Instructions")

Do Until ws.Cells(a, 1).Value = "NewLap"
a = a + 1
Loop
LapRow = a
TypeRow = LapRow + 1
BeforeRow = TypeRow + 1
ReqRow = BeforeRow + 1
AddRow = ReqRow + UBound(RequestedCuts)
AfterRow = AddRow + UBound(RequestedCuts)

oValues(1) = LapRow
oValues(2) = TypeRow
oValues(3) = BeforeRow
oValues(4) = ReqRow
oValues(5) = AddRow
oValues(6) = AfterRow

FindLapRanges = oValues

End Function

Function FindPreCutRanges(PreCuts, ws)
Dim oValues(1 To 4) As Integer, a As Integer
Do Until ws.Cells(a, 1).Value = "NewPreCuts"
 a = a + 1
Loop
oValues(1) = a
oValues(2) = a + 1
oValues(3) = a + 2
oValues(4) = a + 1 + UBound(PreCuts)

FindPreCutRanges = oValues

End Function

Function FindDirectionsRanges(ws)
Dim oValues(1 To 4) As Integer, a As Integer

Do Until ws.Cells(a, 1).Value = "NewDirections"
 a = a + 1
Loop
oValues(1) = a
oValues(2) = a + 1
oValues(3) = a + 2
oValues(4) = a + 3

FindDirectionsRanges = oValues

End Function

Function FindPendingRanges(Pending, ws)
Dim oValues(1 To 4) As Integer, a As Integer

Do Until ws.Cells(a, 1).Value = "NewPending"
a = a + 1
Loop
oValues(1) = a
oValues(2) = a + 1
oValues(3) = a + 2
oValues(4) = a + 1 + UBound(Pending)

FindPendingRanges = oValues

End Function

Sub PopulateRanges(Laps, Mode)
Dim a As Integer, AddCuts() As Integer, RequestedCuts() As Integer, pcRanges As Integer
Dim BeforeSpool() As Integer, AfterSpool() As Integer, PreCuts() As Integer, DirectionsRanges() As Integer
Dim Pending() As Integer, LapRanges() As Integer, WireTypeValues() As Variant, LastCol As Integer, DeSpool() As Integer, srFirstRow As Integer
Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Set ws1 = ThisWorkbook.Worksheets("Instructions")
Set ws3 = ThisWorkbook.Worksheets("Projection")

WireTypeValues = pCollectTotals(WireType)
PreCuts = pCollectTotals(PreCuts)
Pending = pCollectTotals(Pending)

LastCol = UBound(WireTypeValues) + 1

If Mode = 1 Then
 ws1.Cells(1, 1).Value = "NewDirections"
 ws1.Cells(5, 1).Value = "NewLap"
 Set ws2 = ThisWorkbook.Worksheets("Instructions")
End If
If Mode = 2 Then
 ws1.Cells(1, 1).Value = "NewLap"
 ws2.Cells(1, 1).Value = "NewDirections"
 Set ws2 = ThisWorkbook.Worksheets("Instructions2")
End If


For a = 1 To Laps
 RequestedCuts = pCollectLap("LapRequestedCuts", a)
 AddCuts = pCollectLap("AddCuts", a)
 BeforeSpool = pCollectLap("LapBeforeSpool", a)
 AfterSpool = pCollectLap("LapAfterSpool", a)
 DeSpool = FindDeSpool(a)
 LapRanges = FindLapRanges(RequestedCuts, AddCuts)
 DirectionsRanges = FindDirectionsRanges(ws2)

'set standard lap
 Call SetRange(LapRange, LapRange(1), LapRanges(1), 1, LastCol, ws1)
 Call SetRange(WireType, LapRange(2), LapRanges(2), 2, LastCol, ws1)
 Call SetRange(BeforeSpool, LapRange(3), LapRanges(3), 2, LastCol, ws1)
 Call SetRange(PrimaryCuts, LapRange(4), LapRanges(5) + -1, 2, LastCol, ws1)
 Call SetRange(SecondaryCuts, LapRange(5), LapRanges(6) + -1, 2, LastCol, ws1)
 Call SetRange(AfterSpool, LapRange(6), LapRanges(6), 2, LastCol, ws1)
'set side lap
 Call SetRange(SideWireType, LapRange(2), LapRange(2), 1, 1, ws1)
 Call SetRange(SideBefore, LapRange(3), LapRange(3), 1, 1, ws1)
 Call SetRange(SidePrimCuts, LapRange(4), LapRange(5) + -1, 1, 1, ws1)
 Call SetRange(SideSecCuts, LapRange(5), LapRange(6) + -1, 1, 1, ws1)
 Call SetRange(SideAfter, LapRange(6), LapRange(6), 1, 1, ws1)
'set directions range
 Call SetRange(DirLap, DirectionsRanges(1), directoinsranges(1), 1, LastCol, ws2)
 Call SetRange(DirWireType, DirectionsRanges(2), DirectionsRanges(2), 2, LastCol, ws2)
 Call SetRange(DirRemove, DirectionsRanges(3), DirectionsRanges(3), 2, LastCol, ws2)
 Call SetRange(DirAdd, DirectionsRanges(4), DirectionsRanges(4), 2, LastCol, ws2)
'set directions side range
 Call SetRange(DirSideType, DirectionsRanges(2), DirectionsRanges(2), 1, 1, ws2)
 Call SetRange(DirSideRemove, DirectionsRanges(3), DirectionsRanges(3), 1, 1, ws2)
 Call SetRange(DirSideAdd, DirectionsRanges(4), DirectionsRanges(4), 1, 1, ws2)

'color standard lap
 Call BordersColorsFontsValues(LapRange, "Header", 119, 119, 119, 0, 14, "Black", "LAP" & a, 45)
 Call BordersColorsFontsValues(WireType, "Standard", 0, 112, 192, "Top", 14, "Black", WireTypeValues, 30)
 Call BordersColorsFontsValues(BeforeSpool, "Standard", 248, 203, 172, 0, 12, "Black", BeforeSpool, 30)
 Call BordersColorsFontsValues(PrimaryCuts, "Standard", 255, 255, 255, 0, 12, "Black", RequestedCuts, 30)
 Call BordersColorsFontsValues(SecondaryCuts, "Standard", 153, 255, 153, 0, 12, "Black", AddCuts, 30)
 Call BordersColorsFontsValues(AfterSpool, "Standard", 189, 215, 238, "Bottom", 12, "Black", AfterSpool, 30)
'color side lap
 Call BordersColorsFontsValues(SideWireType, "Side", 0, 112, 192, 0, 14, "Black", "Wire Type", 30)
 Call BordersColorsFontsValues(SideBefore, "Side", 248, 203, 172, 0, 14, "Black", "Initial Spool", 30)
 Call BordersColorsFontsValues(SidePrimCuts, "Side", 255, 255, 255, 0, 14, "Black", "Primary Cuts", 30)
 Call BordersColorsFontsValues(SideSecCuts, "Side", 153, 255, 153, 0, 14, "Black", "Secondary Cuts", 30)
 Call BordersColorsFontsValues(SideAfter, "Side", 189, 215, 238, 0, 14, "Black", "Resulting Spool", 30)
'color directions lap
 Call BordersColorsFontsValues(DirLap, "Header", 119, 119, 119, 0, 14, "Black", "LAP" & a, 30)
 Call BordersColorsFontsValues(DirWireType, "Standard", 0, 112, 192, "Top", 14, "Black", WireTypeValues, 30)
 Call BordersColorsFontsValues(DirRemove, "Standard", 208, 206, 206, 0, 14, "Black", DeSpool, 30)
 Call BordersColorsFontsValues(DirAdd, "Standard", 189, 215, 238, 0, 14, "Black", BeforeSpool, 30)
'color directions side lap
 Call BordersColorsFontsValues(DirWireType, "Side", 0, 112, 192, "Top", 14, "Black", WireTypeValues, 30)
 Call BordersColorsFontsValues(DirRemove, "Side", 208, 206, 206, 0, 14, "Black", DeSpool, 30)
 Call BordersColorsFontsValues(DirAdd, "Side", 189, 215, 238, "Bottom", 14, "Black", BeforeSpool, 30)
   
 If Mode = 1 Then
  ws1.Cells(AfterSpool.Rows(AfterSpool.Rows.Count).Row + 2, 1).Value = "NewDirections"
  If a < Laps Then ws1.Cells(AfterSpool.Rows(AfterSpool.Rows.Count).Row + 6, 1).Value = "NewLap"
 End If
 
 If Mode = 2 Then
  If a < Laps Then ws1.Cells(AfterSpool.Rows(AfterSpool.Rows.Count).Row + 2, 1).Value = "NewLap"
 End If
 
 For b = 1 To UBound(DeSpool)
  If DeSpool(b) > "" Then
   Call PrintSpools
  End If
 Next b
  
 
Next a

srFirstRow = ws2.Cells(DirAdd.Row + 2, 1)
Call CopyAndPasteRange(srFirstRow, "WireToSites", ws3, ws2)

End Function

Sub CopyAndPasteRange(FirstRow, mRange, Fws, Tws)
iValues() As Integer

iValues = ImportRange(mRange)
Call SetRange(iRange, iValues(1), iValues(2), iValues(3), iValues(4), Fws)
iRange.Copy Tws.Cells(FirstRow, 1)

End Sub

Sub PrintSpools(WireType, Quanity)
Dim ws As Worksheet, PrintRange As Range
Set ws = ThisWorkbook.Worksheets("PrintSpool")
Set PrintRange = Range(ws.Cells(1, 1), ws.Cells(6, 1))

With PrintRange.Cells(2, 1)
 .Value = WireType
 .Font.Size = 72
End With
With PrintRange.Cells(4, 1)
 .Value = "QTY:" & Quanity
 .Font.Size = 72
End With

End Sub


