Attribute VB_Name = "mProjMisc"
Option Explicit
Dim SiteName As Range, JobName As Range, SiteWireRange As Range, HeaderRange As Range, WireTypeRange As Range
Dim SiteHeader As Range, JobHeader As Range

Sub MatchSites()
Dim Sites() As Variant, SiteWire() As Integer, PreCuts() As Variant
Dim a As Integer, b As Integer, JobNumbers() As Variant, ws As Worksheet, cell As Range
Dim iValues() As Integer, FirstRow As Integer, WireType() As Variant, CurrentRow As Integer, LastRow As Integer
Dim iValues2() As Integer, FromSpool() As Integer, Pending() As Integer, iValues3() As Integer
Set ws = ThisWorkbook.Worksheets("Projection")

Sites = FindSites
JobNumbers = FindJobNumbers
WireType = pCollectTotals("WireType", False)
PreCuts = FindCuts("PreCuts", WireType)
FromSpool = FindCuts("FromSpool", WireType)
Pending FindCuts("Pending", WireType)
LastCol = UBound(WireType)
FirstRow = FindLastRow + 1

Call SetRange(HeaderRange, FirstRow, FirstRow, 1, LastCol)
Call SetRange(WireTypeRange, FirstRow + 1, FirstRow + 1, 3, LastCol)
Call SetRange(JobHeader, FirstRow + 1, FirstRow + 1, 1, 1)
Call SetRange(SiteHeader, FirstRow + 1, FirstRow + 1, 2, 2)

Call BordersColorsFontsValues(HeaderRange, "Header", 119, 119, 119, 0, 28, "Black", "Wire to Sites", 45)
Call BordersColorsFontsValues(JobHeaderRange, "Side", 112, 48, 160, "Top", 14, "Black", "Job#", 30)
Call BordersColorsFontsValues(SiteHeaderRange, "Side", 112, 48, 160, "Top", 14, "Black", "Site ID", 30)
Call BordersColorsFontsValues(WireTypeRange, "wType", 0, 112, 192, 0, 14, "Black", "Site ID", 30)

For a = 1 To UBound(Sites)
 SiteWire = FindSiteWire
 iValues = FindCutTypesForSite(site, PreCuts)
 iValues2 = FindCutTypesForSite(site, FromSpool)
 iValues3 = FindCutTypesForSite(site, Pending)
 CurrentRow = FirstRow + 1 + a

 Call SetRange(JobName, CurrentRow, CurrentRow, 1, 1, ws)
 Call SetRange(SiteName, CurrentRow, CurrentRow, 2, 2, ws)
 Call SetRange(SiteWireRange, CurrentRow, CurrentRow, 3, LastCol, ws)
 
 Call BordersColorsFontsValues(JobName, "Standard", 112, 48, 160, 0, 14, "Black", JobNumbers(a), 30)
 Call BordersColorsFontsValues(SiteName, "Side", 112, 48, 160, 0, 14, "Black", Sites(a), 30)
 Call BordersColorsFontsValues(SiteWireRange, "Standard", 255, 255, 255, 0, 14, "Black", "", 30)
 
 For b = 1 To UBound(iValues)
  If iValues(b) > "" Then
   SiteWireRange.Cells(1, b).Value = iValues(b)
   SiteWireRange.Cells(1, b).Interior.Color = RGB(146, 208, 80)
  End If
 Next b
 For b = 1 To UBound(iValues2)
  If iValues2(b) > "" Then
   SiteWireRange.Cells(1, b).Value = iValues2(b)
   SiteWireRange.Cells(1, b).Interior.Color = RGB(0, 176, 80)
  End If
 Next b
 For b = 1 To UBound(iValues3)
  If iValues3(b) > "" Then
   SiteWireRange.Cells(1, b).Value = iValues3(b)
   SiteWireRange.Cells(1, b).Interior.Color = RGB(255, 80, 80)
  End If
 Next b
 
Next a

LastRow = SiteWireRange.Rows(SiteWireRange.Rows.Count).Row + 1
ws.Cells(LastRow, 1).Value = "LastRow"
ws.Cells(LastRow, 1).Font.ColorIndex = -4142
ws.Cells(LastRow, 2).Value = Now

End Sub

Function FindCuts(mType, WireType)
Dim oValues() As Variant, Wire As Variant, iValues() As Integer
Dim a As Integer, b As Integer

For Each Wire In WireType
iValues = pCollectTotals(mType, Wire)
b = 0
 Do
  a = a + 1
  b = b + 1
  ReDim Preserve oValues(1 To a)
  oValues(a) = iValues(b)
 Loop
a = a + 1
ReDim Preserve oValues(1 To a)
oValues(a) = "end"
Next Wire

FindCuts = oValues

End Function

Function FindJobNumbers()
Dim ws As Worksheet, a As Integer, oValues() As Variant
Set ws = ThisWorkbook.Worksheets("Calculate")

Do
a = a + 1
If ws.Cells(a, 2).Value = "Job#" Then
 Do Until ws.Cells(a + 1, 2).Value = ""
  a = a + 1
  b = b + 1
  ReDim Preserve oValues(1 To b)
  oValues(b) = ws.Cells(a, 2).Value
 Loop
End If
Loop

FindJobNumber = oValues

End Function


Function FindSites()
Dim ws As Worksheet, a As Integer, oValues() As Variant
Set ws = ThisWorkbook.Worksheets("Calculate")

Do
a = a + 1
If ws.Cells(a, 2).Value = "Site ID" Then
 Do Until ws.Cells(a + 1, 2).Value = ""
  a = a + 1
  b = b + 1
  ReDim Preserve oValues(1 To b)
  oValues(b) = ws.Cells(a, 2).Value
 Loop
End If
Loop

FindSites = oValues

End Function

Function FindSiteWire(SiteName, wirename)
Dim ws As Worksheet, a As Integer, oValues() As Variant, c As Integer, NameRange As Integer
Set ws = ThisWorkbook.Worksheets("Calculate")

Do
a = a + 1
If ws.Cells(a, 2).Value = SiteName Then
NameRange = a
Exit Do
Loop

a = 2
Do
a = a + 1
If ws.Cells(7, a).Value = "" Then Exit Do
c = c + 1
ReDim Preserve oValues(1 To c)
oValues(c) = a
Loop

FindSiteWire = oValues

End Function

Function FindCutTypesForSite(site, ByRef Cuts As Variant)
Dim a As Integer, b As Integer
Dim oValues() As Variant, SiteWire() As Integer

ReDim oValues(1 To UBound(WireType))
SiteWire = FindSiteWire

For Each Wire In SiteWire
 Do Until Cuts(a) = "end"
  a = a + 1
  If Wire = Cuts(a) Then
   b = b + 1
   ReDim Preserve oValues(1 To b)
   oValues = Cuts(a)
   Cuts(a) = "Selected"
  End If
 Loop
Next Wire

FindCutTypesForSite = oValues

End Function

Function FindLastRow()
Dim ws As Worksheet, a As Integer
Set ws = ThisWorkbook.Worksheets("Projection")

Do
a = a + 1
If ws.Cells(a, 1) = "TOTALS" Then
 Do
  a = a + 1
  If ws.Cells(a + 1, 1).Interior.ColorIndex = -4142 Then
   FindLastRow = a
   Exit Function
  End If
 Loop
Loop

End Function

Sub SetRange(mRange, FirstRow, LastRow, FirstCol, LastCol, ws)
Set mRange = Range(ws.Cells(FirstRow, LastRow), ws.Cells(LastRow, LastCol))
End Sub
