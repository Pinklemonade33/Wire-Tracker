VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub DoneButton_Click()
Cells(10, 1).Value = True
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim a As Integer, b As Integer
Dim FirstCutsRange As Integer, LastCutsRange As Integer, FirstSpoolRange As Integer, LastSpoolRange As Integer
Dim FirstPreCutsRange As Integer, LastPreCutsRange As Integer, LastRow As Integer
Dim FirstSelectedPreCutsRange As Integer, LastSelectedSpoolRange As Integer, LastSelectedPreCutsRange As Integer


Do Until Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
 With Cells(a, 1)
  If .Value = "Requested Cuts" Then FirstCutsRange = a
  If .Value = "Spools" Then
   FirstSpoolRange = a + 1
   LastCutsRange = a + -1
  End If
  If .Value = "Pre-Cuts" Then
   FirstPreCutsRange = a + 1
   LastSpoolRange = a + -1
  End If
 End With
 If Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then LastPreCutsRange = a + 1
Loop

a = 0
Do Until Cells(1, a + 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
 If Cells(1, a).Value = "Selected Pre-Cuts" Then FirstSelectedPreCutsRange = a
 If Cells(1, a).Value = "Actions" Then LastSelectedPreCutsRange = a + -2
Loop
a = 0
LastSelectedSpoolRange = FirstSelectedPreCutsRange + -2

Do Until Cells(a + 1, 1).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
 a = a + 1
Loop
LastRow = a

a = ActiveCell.Row
b = ActiveCell.Column
With Cells(a, b)
 If b < 7 Then
  If a >= FirstCutsRange And a <= LastCutsRange Then
   If .Interior.ColorIndex = 4 Then
     Call UnPick(0, 255, 0, FirstPreCutsRange, LastPreCutsRange, 1, 6, a, b, 2, LastRow, FirstSelectedPreCutsRange, LastSelectedPreCutsRange, False, False)
     Exit Sub
   End If
   If .Interior.Color = RGB(255, 255, 255) Then
     Call Pick(0, 176, 80, FirstCutsRange, LastCutsRange, a, b, 5, LastRow, 8, 8, True)
     Exit Sub
   End If
  End If
  If a >= FirstSpoolRange And a <= LastSpoolRange Then
   If .Interior.Color = RGB(255, 255, 255) And .Value > "" Then
     Call PickSpool(1, 6, a, b, LastRow)
     Exit Sub
   End If
  End If
  If a >= FirstPreCutsRange And a <= LastPreCutsRange Then
   If .Interior.Color = RGB(255, 255, 255) Then
     Call Pick(0, 255, 0, FirstCutsRange, LastCutsRange, a, b, 2, LastRow, FirstSelectedPreCutsRange, LastSelectedPreCutsRange, False)
     Exit Sub
   End If
   If .Interior.ColorIndex = 4 Then
     Call UnPick(0, 255, 0, FirstPreCutsRange, LastPreCutsRange, 1, 6, a, b, 2, LastRow, FirstSelectedPreCutsRange, LastSelectedPreCutsRange, False, False)
     Exit Sub
   End If
    
 End If
End If
 If b > 9 And b <= LastSelectedSpoolRange Then
  If a = 3 Then
    Call SelectSpool(b, 10, LastSelectedSpoolRange, LastRow, False)
    Exit Sub
  End If
 End If
 If b = 8 Then
  If a = 3 Then
    Call UnpickSpool(FirstSpoolRange, LastSpoolRange, LastRow, 9, LastSelectedSpoolRange, FirstCutsRange, LastCutsRange)
    Exit Sub
  End If
  If a > 4 And a <= LastRow Then
    Call UnPick(0, 176, 80, FirstCutsRange, LastCutsRange, 1, 6, a, b, FirstPreCutsRange, LastPreCutsRange, 1, 6, True, True)
    Exit Sub
  End If
 End If
 If b >= FirstSelectedPreCutsRange And b <= LastSelectedPreCutsRange Then
  If a >= 2 And a <= LastRow Then Call UnPick(0, 255, 0, FirstCutsRange, LastCutsRange, 1, 6, a, b, FirstPreCutsRange, LastPreCutsRange, 1, 6, False, True)
 End If
End With
   
End Sub




