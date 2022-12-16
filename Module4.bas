Attribute VB_Name = "Module4"
Option Explicit
Dim swire() As Variant, inv As Worksheet, sett As Worksheet, i As Integer

Sub AddWire()
Dim a As Integer, b As Integer, c As Integer, LastRange As Integer, lastrange2 As Integer, same As Boolean, size1 As Integer, blankcheck As Boolean
Dim r1 As Range, r2 As Range, ws1 As Worksheet, cell As Range, ws2 As Worksheet, sarray() As Variant, math As Integer, d As Integer, firstrange2 As Integer

Set ws2 = ThisWorkbook.Worksheets("saved")
Set ws1 = ThisWorkbook.Worksheets("settings")

b = 2
Do Until b = 6
    For a = 3 To 20
        If ws1.Cells(a, b).Value = "Mode" Then a = a + 3
        If ws1.Cells(a, b).Value > "" And Not blankcheck = True Then c = a + 1
        If ws1.Cells(a, b).Value = "" Then blankcheck = True
         If blankcheck = True And ws1.Cells(a, b).Value > "" Then
           ws1.Cells(c, b).Value = ws1.Cells(a, b).Value
           ws1.Cells(a, b).Value = ""
           c = c + 1
         End If
    Next a
c = 0
b = b + 1
blankcheck = False
Loop
    
    
a = 3
b = 1
 Do
  a = a + 1
   If a > LastRange Then LastRange = a
    If ws1.Cells(a, b).Value = "" Then
     b = b + 1
     a = 3
    End If
  If b = 6 Then Exit Do
 Loop
        LastRange = LastRange + -1
          Set r1 = Range(ws1.Cells(2, 1), ws1.Cells(LastRange, b))
            size1 = LastRange + -1
              ReDim sarray(1 To (size1 * b))
     
     
a = 0
Do
 a = a + 1
  If ws1.Cells(2, 1).Value = ws2.Cells(a, 1).Value Then same = True
  If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then Exit Do
Loop
        firstrange2 = a
          lastrange2 = firstrange2 + size1
            Set r2 = Range(ws2.Cells(firstrange2, 1), ws2.Cells(lastrange2 + -1, b))

If same = False Then
    

For d = a To lastrange2
    ws2.Cells(d, 1).Interior.ColorIndex = 3
Next d

c = 1
For Each cell In r1
    sarray(c) = cell.Value
    c = c + 1
Next cell

c = 1
For Each cell In r2
   cell.Value = sarray(c)
       c = c + 1
Next cell

ws2.Cells(lastrange2, 1).Value = "end"

End If

If same = True Then MsgBox ("The same wire type has already been created, if you would like to make changes please click the edit button")

End Sub

Sub RemoveWire()
Dim a As Integer, b As Variant, c As Integer, LastRange As Integer, FirstRange As Integer, shiftarr() As Variant, done As Boolean, firstrange2 As Integer, d As Integer
Dim r2 As Range, ws2 As Worksheet, ws1 As Worksheet, cell As Range, r3 As Range, lastrange2 As Integer, r4 As Range, size4 As Integer, lastrange3 As Integer, wirename As Integer

Set ws1 = ThisWorkbook.Worksheets("settings")
Set ws2 = ThisWorkbook.Worksheets("saved")


Do
 a = a + 1
  If ws2.Cells(a, 1).Value = ws1.Cells(2, 1).Value Then Exit Do
    If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then
       done = True
       Exit Do
    End If
Loop
        FirstRange = a


If Not done = True Then



Do
a = a + 1
  If ws2.Cells(a, 1).Value = "end" Then Exit Do
Loop
        LastRange = a
        Set r2 = Range(ws2.Cells(FirstRange, 1), ws2.Cells(LastRange, 5))
        
        
r2.Delete (xlShiftUp)

  
For a = 0 To ws1.OLEObjects("combobox1").object.ListCount + -1
    If ws1.Cells(2, 1).Value = ws1.OLEObjects("combobox1").object.list(a) Then
        wirename = a
                ws1.OLEObjects("combobox1").object.RemoveItem wirename
        Exit For
    End If
Next a



End If


If done = True Then MsgBox ("This wire does not exist")
    
End Sub


Sub selectwire()
Dim a As Integer, b As Integer, c As Integer, FirstRange As Integer, LastRange As Integer, done As Boolean, rr1 As Range, rr11 As Range, noselection As Boolean
Dim cell As Range, ws1 As Worksheet, ws2 As Worksheet, selectarr() As Variant, r1 As Range, r2 As Range, size1 As Integer, missingitem As Variant, comboval As Variant

Set ws1 = ThisWorkbook.Worksheets("settings")
Set ws2 = ThisWorkbook.Worksheets("saved")
Set rr1 = Range(ws1.Cells(4, 2), ws1.Cells(21, 5))
comboval = ws1.OLEObjects("combobox1").object.Value
If ws1.OLEObjects("combobox1").object.Value = "" Then done = True


Do
  a = a + 1
  If ws2.Cells(a, 1).Value = comboval Then Exit Do
  If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then
    done = True
    Exit Do
  End If
Loop
        FirstRange = a


If Not done = True Then



For Each cell In rr1
    cell.Value = ""
Next cell

    
Do
a = a + 1
  If ws2.Cells(a, 1).Value = "end" Then Exit Do
Loop
        LastRange = a + -1
          Set r1 = Range(ws2.Cells(FirstRange, 1), ws2.Cells(LastRange, 6))
            size1 = LastRange + -(FirstRange + -1)
              Set r2 = Range(ws1.Cells(2, 1), ws1.Cells(1 + size1, 6))
              ReDim selectarr(1 To (size1 * 6))


For Each cell In r1
    c = c + 1
    selectarr(c) = cell.Value
Next cell
c = 0
For Each cell In r2
    c = c + 1
    cell.Value = selectarr(c)
Next cell
End If


If done = True Then
    MsgBox ("This wire does not exist")
        If ws1.OLEObjects("combobox1").object.Value > "" Then
            For a = 0 To ws1.OLEObjects("combobox1").object.ListCount + -1
                If ws1.OLEObjects("combobox1").object.Value = ws1.OLEObjects("combobox1").object.list(a) Then
                    missingitem = a
                    ws1.OLEObjects("combobox1").object.RemoveItem missingitem
                End If
            Next a
        End If
End If


ws1.OLEObjects("combobox1").object.Value = ""

End Sub

Sub editwire()
Dim a As Integer, b As Integer, c As Integer, lastrange1 As Integer, size1 As Integer, size2 As Integer, r4 As Range, verylastrange2 As Integer, firstrange2 As Integer, size3 As Integer, done As Boolean
Dim ws1 As Worksheet, ws2 As Worksheet, cell As Range, r1 As Range, r2 As Range, lastrange2 As Integer, editarr() As Variant, r3 As Range, verylastrange As Integer, editarr2() As Variant, sizediff As Integer
Dim lastrange3 As Integer, size4 As Integer, blankcheck As Boolean

Set ws1 = ThisWorkbook.Worksheets("settings")
Set ws2 = ThisWorkbook.Worksheets("saved")


b = 2
Do Until b = 6
 For a = 3 To 20
    If ws1.Cells(a, b).Value = "Mode" Then a = a + 3
    If ws1.Cells(a, b).Value > "" And Not blankcheck = True Then c = a + 1
    If ws1.Cells(a, b).Value = "" Then blankcheck = True
    
    If blankcheck = True And ws1.Cells(a, b).Value > "" Then
       ws1.Cells(c, b).Value = ws1.Cells(a, b).Value
       ws1.Cells(a, b).Value = ""
       c = c + 1
    End If
 Next a
c = 0
b = b + 1
blankcheck = False
Loop


a = 0
Do
a = a + 1
  If ws2.Cells(a, 1).Value = ws1.Cells(2, 1).Value Then Exit Do
  If ws2.Cells(a, 1).Interior.ColorIndex = -4142 Then
    done = True
    Exit Do
  End If
Loop
        firstrange2 = a

If Not done = True Then
    
   
Do
a = a + 1
  If ws2.Cells(a, 1).Value = "end" Then Exit Do
Loop
        lastrange2 = a
          Set r2 = Range(ws2.Cells(firstrange2, 1), ws2.Cells(lastrange2, b))
          size2 = lastrange2 + -(firstrange2 + -1)

        
a = 3
b = 1
Do
a = a + 1
  If ws1.Cells(a, b).Value = "" Then
    b = b + 1
    a = 3
      
     Else
     If a > lastrange1 Then lastrange1 = a
  End If
If b = 6 Then Exit Do
Loop
        Set r1 = Range(ws1.Cells(2, 1), ws1.Cells(lastrange1 + 1, b))
        size1 = lastrange1
          ReDim editarr(1 To (size1 * b))
          size3 = size1 + -size2
            verylastrange = (lastrange2 + size3) + 1
              Set r3 = Range(ws2.Cells(firstrange2, 1), ws2.Cells(verylastrange, b))
       
c = 0
For Each cell In r1
    c = c + 1
    editarr(c) = cell.Value
Next cell

If size1 = size2 Then
c = 0
  For Each cell In r2
    c = c + 1
    cell.Value = editarr(c)
  Next cell
End If

ws2.Cells(lastrange2, 1).Value = "end"
    
    
If size1 <> size2 Then

r2.Delete (xlShiftUp)

c = 0
    For Each cell In r3
        c = c + 1
        cell.Value = editarr(c)
    Next cell
    
    For a = firstrange2 To verylastrange
        ws2.Cells(a, 1).Interior.ColorIndex = 3
    Next a

ws2.Cells(verylastrange, 1).Value = "end"

End If
    
    


End If

If done = True Then MsgBox ("This wire does not exist")

End Sub

Sub RefreshWireListSettings()
Dim a As Integer, b As Integer, c As Integer
Dim ws1 As Worksheet, ws2 As Worksheet

Set ws1 = ThisWorkbook.Worksheets("settings")
Set ws2 = ThisWorkbook.Worksheets("saved")

For a = 1 To ws1.OLEObjects("combobox1").object.ListCount
ws1.OLEObjects("combobox1").object.RemoveItem 0
Next a

a = 1
 Do Until ws2.Cells(a + 1, 1).Value = ""
  ws1.OLEObjects("combobox1").object.AddItem ws2.Cells(a, 1).Value
   Do Until ws2.Cells(a, 1).Value = "end"
            a = a + 1
   Loop
            a = a + 1
 Loop

End Sub


