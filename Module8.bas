Attribute VB_Name = "Module8"
Option Explicit

Sub UnPick(Red, Green, Blue, FirstRow, LastRow, FirstCol, LastCol, ActiveRow, ActiveCol, SelFirstRow, SelLastRow, SelFirstCol, SelLastCol, ForSpool, FromSel)
Dim MN As Worksheet, cell As Range
Set MN = ThisWorkbook.Worksheets("Manual")


For Each cell In Range(MN.Cells(FirstRow, 1), MN.Cells(LastRow, 6))
 If cell.Value = MN.Cells(ActiveRow, ActiveCol).Value And cell.Interior.Color = RGB(Red, Green, Blue) Then
  cell.Interior.Color = RGB(255, 255, 255)
  Exit For
 End If
Next cell


If ForSpool = False Then
For Each cell In Range(MN.Cells(SelFirstRow, SelFirstCol), MN.Cells(SelLastRow, SelLastCol))
 If cell.Value = MN.Cells(ActiveRow, ActiveCol).Value Then
  cell.Interior.Color = RGB(255, 255, 255)
  If FromSel = False Then cell.Value = ""
  Exit For
 End If
Next cell
End If

If ForSpool = True Then MN.Cells(4, 8).Value = MN.Cells(4, 8).Value + MN.Cells(ActiveRow, ActiveCol).Value

MN.Cells(ActiveRow, ActiveCol).Interior.Color = RGB(255, 255, 255)
If FromSel = True Then MN.Cells(ActiveRow, ActiveCol).Value = ""

End Sub

Sub Pick(Red, Green, Blue, FirstCutRange, LastCutRange, ActiveRow, ActiveCol, SelFirstRow, SelLastRow, SelFirstCol, SelLastCol, ForSpool)
Dim MN As Worksheet, cell As Range, good As Boolean
Set MN = ThisWorkbook.Worksheets("Manual")

good = True
If ForSpool = True Then
 If MN.Cells(3, 8).Value = "" Then good = False
End If

If good = True Then
good = False
For Each cell In Range(MN.Cells(FirstCutRange, 1), MN.Cells(LastCutRange, 6))
 If cell.Value = MN.Cells(ActiveRow, ActiveCol).Value And cell.Interior.Color = RGB(255, 255, 255) Then
  good = True
  If ForSpool = False Then cell.Interior.Color = RGB(Red, Green, Blue)
  Exit For
 End If
Next cell
End If

If good = True Then

  good = False
    For Each cell In Range(MN.Cells(SelFirstRow, SelFirstCol), MN.Cells(SelLastRow, SelLastCol))
      If cell.Value = "" Then good = True
    Next cell
  If good = False Then
    Call ExtendRows(SelLastRow)
    SelFirstRow = SelFirstRow + 1
  End If
   
    For Each cell In Range(MN.Cells(SelFirstRow, SelFirstCol), MN.Cells(SelLastRow, SelLastCol))
     If cell.Value = "" Then
      cell.Interior.Color = RGB(Red, Green, Blue)
      cell.Value = MN.Cells(ActiveRow, ActiveCol).Value
      Exit For
     End If
    Next cell
    
    If ForSpool = True Then
     If MN.Cells(4, 8).Value = "" Then MN.Cells(4, 8).Value = MN.Cells(3, 8).Value
    End If
    
    If ForSpool = True Then MN.Cells(4, 8).Value = MN.Cells(4, 8).Value + -MN.Cells(ActiveRow, ActiveCol).Value
    
    MN.Cells(ActiveRow, ActiveCol).Interior.Color = RGB(Red, Green, Blue)
    
End If
End Sub

Sub PickSpool(FirstCol, LastCol, SelectedRow, SelectedCol, LastRow)
Dim MN As Worksheet, cell As Range, good As Boolean
Set MN = ThisWorkbook.Worksheets("Manual")

good = False
For Each cell In Range(MN.Cells(3, FirstCol), MN.Cells(3, LastCol))
 If cell.Value = "" Then
   good = True
 End If
Next cell

If MN.Cells(3, 8).Value > "" Then Call SelectSpool(8, FirstCol, LastCol, LastRow, True)

MN.Cells(SelectedRow, SelectedCol).Interior.Color = RGB(0, 176, 240)
MN.Cells(3, 8).Value = MN.Cells(SelectedRow, SelectedCol)

End Sub

Sub UnpickSpool(FirstSpoolRow, LastSpoolRow, LastRow, FirstCol, LastCol, FirstCutRange, LastCutRange)
Dim MN As Worksheet, cell As Range, good As Boolean, answer As Boolean, cut As Range
Set MN = ThisWorkbook.Worksheets("Manual")

good = False
For Each cell In Range(MN.Cells(5, 8), MN.Cells(LastRow, 8))
 If cell.Value > "" Then good = True
Next cell
If good = True Then answer = MsgBox("This spool has cuts assigned to it, are you sure you want to remove it?", vbYesNo)

If answer = True Or good = False Then
    For Each cell In Range(MN.Cells(FirstSpoolRow, 1), MN.Cells(LastSpoolRow, 6))
     If cell.Value = MN.Cells(3, 8).Value Then
      cell.Interior.Color = RGB(255, 255, 255)
      Exit For
     End If
    Next cell
    
    Range(MN.Cells(3, 8), MN.Cells(4, 8)).Value = ""
    
    For Each cell In Range(MN.Cells(5, 8), MN.Cells(LastRow, 8))
     For Each cut In Range(MN.Cells(FirstCutRange, 1), MN.Cells(LastCutRange, 6))
      If cut.Value = cell.Value And cut.Interior.Color = RGB(0, 176, 80) Then cut.Interior.Color = RGB(255, 255, 255)
     Next cut
     cell.Value = ""
     cell.Interior.Color = RGB(255, 255, 255)
    Next cell

    good = True
    For Each cell In Range(MN.Cells(3, FirstCol), MN.Cells(3, LastCol))
     If cell.Value = "" Then good = False
    Next cell
    
   If good = True Then Call ExtendSpools(False)
End If
End Sub


Sub SelectSpool(ActiveCol, FirstCol, LastCol, LastRow, FromSpools)
Dim MN As Worksheet, cell As Range, Active() As Integer, inactive() As Integer
Dim a As Integer, InactiveCol As Integer, LastSpoolsRange As Integer
Set MN = ThisWorkbook.Worksheets("Manual")

a = 10
Do Until MN.Cells(3, a + 1).Interior.Color = RGB(119, 119, 119)
 a = a + 1
Loop
LastSpoolsRange = a


ReDim Active(1 To -3 + (LastRow + 1))
ReDim inactive(1 To -3 + (LastRow + 1)) As Integer


If FromSpools = False Then
    a = 0
    For Each cell In Range(MN.Cells(3, 8), MN.Cells(LastRow, 8))
     If cell.Value > "" Then
      a = a + 1
      Active(a) = cell.Value
      cell.Value = ""
     End If
    Next cell
    a = 0
    For Each cell In Range(MN.Cells(3, ActiveCol), MN.Cells(LastRow, ActiveCol))
     If cell.Value > "" Then
      a = a + 1
      inactive(a) = cell.Value
      cell.Value = ""
     End If
    Next cell
    a = 0
    For Each cell In Range(MN.Cells(3, ActiveCol), MN.Cells(LastRow, ActiveCol))
     If Active(a + 1) > 0 Then
      a = a + 1
      cell.Value = Active(a)
     End If
    Next cell
    a = 0
    For Each cell In Range(MN.Cells(3, 8), MN.Cells(LastRow, 8))
     If inactive(a + 1) > 0 Then
      a = a + 1
      cell.Value = inactive(a)
     End If
    Next cell
End If


If FromSpools = True Then
    a = 0
    For Each cell In Range(MN.Cells(3, 8), MN.Cells(LastRow, 8))
     If cell.Value > "" Then
      a = a + 1
      Active(a) = cell.Value
     End If
    Next cell
    
    For a = 10 To LastSpoolsRange
     If MN.Cells(3, a).Value = "" Then
      InactiveCol = a
      Exit For
     End If
    Next a
    a = 0
    For Each cell In Range(MN.Cells(3, InactiveCol), MN.Cells(LastRow, InactiveCol))
     If Active(a + 1) > 0 Then
      a = a + 1
      cell.Value = Active(a)
     End If
    Next cell
End If

End Sub

Sub ExtendRows(NewRow)
Dim a As Integer, BorderRanges(1 To 5) As Integer, LastCol As Integer, b As Integer

NewRow = NewRow + 1
 
a = 1
Do Until Cells(NewRow + -1, a).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
 If Cells(NewRow + -1, a).Interior.Color = RGB(119, 119, 119) Then
  b = b + 1
  BorderRanges(b) = a
 End If
Loop
LastCol = a + -1

Call ManualColors(NewRow, NewRow, 255, 255, 255, 14, False, "", 30, False, LastCol, 1, False)
Call ManualBorder(BorderRanges(1), NewRow, 3, 1)
Call ManualBorder(BorderRanges(2), NewRow, 1, 2)
Call ManualBorder(BorderRanges(3), NewRow, 1, 1)
Call ManualBorder(BorderRanges(4), NewRow, 3, 1)
Call ManualBorder(BorderRanges(5), NewRow, 1, 3)

End Sub

Sub PopulateList(WireType, KeyWord, ListName)
Dim a As Integer, ws1 As Worksheet
Dim FirstRow As Integer, LastRow As Integer, cell As Range

Set ws1 = ThisWorkbook.Worksheets("saved")


Do
a = a + 1
If ws1.Cells(a, 1) = WireType Then Exit Do
Loop
    FirstRow = a
Do
a = a + 1
If ws1.Cells(a, 1) = "end" Then Exit Do
Loop
    LastRow = a

For Each cell In Range(ws1.Cells(FirstRow, 1), ws1.Cells(LastRow, 5))
 If cell.Value = KeyWord Then
  a = cell.Row
  Do Until ws1.Cells(a + 1, cell.Column).Value = ""
   a = a + 1
   ListName.object.AddItem ws1.Cells(a, cell.Column).Value
  Loop
 End If
Next cell

End Sub


