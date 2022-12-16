Attribute VB_Name = "Module5"
Option Explicit

Sub RefreshWireList(CboxName, Worksheet)
Dim ws2 As Worksheet, a As Integer, b As Integer, c As Integer
Dim ws1 As Worksheet


Set ws1 = ThisWorkbook.Worksheets(Worksheet)
Set ws2 = ThisWorkbook.Worksheets("saved")

Do Until ws1.OLEObjects(CboxName).object.ListCount = 0
    ws1.OLEObjects(CboxName).object.RemoveItem 0
Loop

 ws1.OLEObjects(CboxName).object.ListRows = 0


a = 0
Do
a = a + 1
  If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then Exit Do
  If ws2.Cells(a, 1).Value > "" Then
   With ws1.OLEObjects(CboxName).object
    .AddItem ws2.Cells(a, 1).Value
    .ListRows = .ListRows + 1
   End With
  End If

    Do Until ws2.Cells(a, 1).Value = "end"
    a = a + 1
    Loop
    
Loop

End Sub

Sub RefreshGroupList()
Dim ws1 As Worksheet, ws2 As Worksheet, a As Integer, b As Integer, c As Integer

Set ws1 = ThisWorkbook.Worksheets("Calculate")
Set ws2 = ThisWorkbook.Worksheets("saved")

Do Until ws1.OLEObjects("ComboBox3").object.ListCount = 0
    ws1.OLEObjects("ComboBox3").object.RemoveItem 0
Loop

a = 0
Do
a = a + 1
  If ws2.Cells(a, 9).Interior.ColorIndex = -4142 Then Exit Do
  If ws2.Cells(a, 9).Value > "" Then ws1.OLEObjects("ComboBox3").object.AddItem ws2.Cells(a, 9).Value
    
    Do Until ws2.Cells(a, 9).Value = "end"
      a = a + 1
    Loop

Loop

End Sub


Sub CreateGroup1()
Dim ws1 As Worksheet, ws2 As Worksheet, a As Integer, b As Integer, c As Integer, firstrange2 As Integer, done As Boolean, lastrange2 As Integer, cell As Range, r2 As Range

Set ws1 = ThisWorkbook.Worksheets("Calculate")
Set ws2 = ThisWorkbook.Worksheets("saved")
If ws1.OLEObjects("combobox3").object.Value = "" Then done = True

If Not done = True Then

Do
a = a + 1
If Cells(a, 9).Interior.ColorIndex = -4142 Then Exit Do
Loop
        ws2.Cells(a, 9).Value = ws1.OLEObjects("combobox3").object.Value
        firstrange2 = a
      

b = firstrange2
a = 2
Do
a = a + 1
b = b + 1
If ws1.Cells(5, a).Value > "" Then
  ws2.Cells(b, 9).Value = ws1.Cells(5, a).Value
  Else: Exit Do
End If
Loop
        lastrange2 = b
         Set r2 = Range(ws2.Cells(firstrange2, 9), ws2.Cells(lastrange2, 9))
         ws2.Cells(lastrange2, 9).Value = "end"
         
For Each cell In r2
    cell.Interior.ColorIndex = 3
Next cell

ws1.OLEObjects("combobox3").object.AddItem ws1.OLEObjects("combobox3").object.Value

End If

If done = True Then MsgBox ("Please fill out the group name inside the dropbox")

End Sub

Sub DeleteGroup1()

Dim ws1 As Worksheet, ws2 As Worksheet, a As Integer, b As Integer, c As Integer, firstrange2 As Integer, done As Boolean, lastrange2 As Integer
Dim lastrange1 As Integer, r1 As Range, r2 As Range, cell As Range

Set ws1 = ThisWorkbook.Worksheets("Calculate")
Set ws2 = ThisWorkbook.Worksheets("saved")
If ws1.OLEObjects("combobox3").object.Value = "" Then done = True

If Not done = True Then



Do
a = a + 1
If ws2.Cells(a, 9).Value = ws1.OLEObjects("combobox3").object.Value Then Exit Do
If ws2.Cells(a, 9).Interior.ColorIndex = -4142 Then Exit Sub
Loop
        firstrange2 = a
        
Do
a = a + 1
If ws2.Cells(a, 9).Value = "end" Then Exit Do
Loop
        lastrange2 = a
          Set r2 = Range(ws2.Cells(firstrange2, 9), ws2.Cells(lastrange2, 9))
          
r2.Delete (xlShiftUp)



a = 0
Do
a = a + 1
If ws1.Cells(5, a).Value = "" Then Exit Do
Loop
        lastrange1 = a
         Set r1 = Range(ws1.Cells(5, 3), ws1.Cells(5, lastrange1))
         
For Each cell In r1
 With cell
  .Value = ""
  .Interior.ColorIndex = -4142
 End With
Next cell

End If

If done = True Then MsgBox ("Please fill out the group name inside the dropbox")

End Sub

Sub SelectGroup()
Dim ws1 As Worksheet, ws2 As Worksheet, a As Integer, b As Integer, c As Integer, firstrange2 As Integer, done As Boolean, lastrange2 As Integer
Dim lastrange1 As Integer, r1 As Range, r2 As Range, cell As Range, firstrange1 As Integer

Set ws1 = ThisWorkbook.Worksheets("Calculate")
Set ws2 = ThisWorkbook.Worksheets("saved")


Do
a = a + 1
 If ws1.Cells(7, a).Value = "" Then Exit Do
Loop
        firstrange1 = a

a = 0
Do
a = a + 1
 If ws2.Cells(a, 9).Interior.ColorIndex = -4142 Then
    MsgBox ("This group does not exist")
    Exit Sub
 End If
 If ws2.Cells(a, 9).Value = ws1.OLEObjects("combobox3").object.Value Then Exit Do
Loop
        firstrange2 = a + 1

Do
a = a + 1
 If ws2.Cells(a, 9).Value = "end" Then Exit Do
Loop
        lastrange2 = a + -1
            Set r2 = Range(ws2.Cells(firstrange2, 9), ws2.Cells(lastrange2, 9))


a = firstrange1
For Each cell In r2
  With Cells(7, a)
   .Value = cell.Value
   .Interior.Color = Sheets("colors").Cells(1, 3).Interior.Color
   .Font.Size = 14
   .Font.Color = Cells(1, 1).Font.Color
  End With
a = a + 1
Next cell

ws1.OLEObjects("combobox3").object.Value = ""

End Sub
         
Sub ClearProjection()
Dim a As Integer, b As Integer, c As Integer
Dim ws1 As Worksheet, ws2 As Worksheet, Rws2 As Range
Dim LastRowWS2 As Integer, LastColWS2 As Integer

Set ws1 = ThisWorkbook.Worksheets("Calculate")
Set ws2 = ThisWorkbook.Worksheets("Projection")


Do
a = a + 1
If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then Exit Do
Loop

                    LastRowWS2 = a

Do
a = a + 1
If ws2.Cells(1, a).Interior.ColorIndex = -4142 Then Exit Do
Loop
                   LastColWS2 = a
                   
                   Set Rws2 = Range(ws2.Cells(1, 1), ws2.Cells(LastRowWS2, LastColWS2))
                   
For Each cell In Rws2
 With cell
   .Value = ""
   .Interior.ColorIndex = -4142
 End With
Next cell

End Sub


