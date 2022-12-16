VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then Cells(8, 1).Value = "LaunchOnStart"
If CheckBox1.Value = False Then Cells(8, 1).Value = ""
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then Cells(9, 1).Value = "LaunchOnCalc"
If CheckBox2.Value = False Then Cells(9, 1).Value = ""
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then Cells(10, 1).Value = "LaunchOnCut"
If CheckBox3.Value = False Then Cells(10, 1).Value = ""
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.Value = True Then Cells(11, 1).Value = "LaunchOnTrim"
If CheckBox4.Value = False Then Cells(11, 1).Value = ""
End Sub

Private Sub ComboBox1_Click()
    Call selectwire
    ComboBox1.Value = ""
End Sub

Private Sub CommandButton1_Click()
    Call selectwire
End Sub


Private Sub DeleteButton_Click()
    Call RemoveWire
End Sub

Private Sub Free_Click()
    Cells(6, 1).Value = "Free"
    Cells(5, 1).Value = 0
    Cells(4, 1).Value = 0
End Sub

Private Sub OptionButton3_Click()
    Cells(8, 1).Value = "Launch On Start"
    Cells(9, 1).Value = 0
End Sub

Private Sub OptionButton4_Click()
    Cells(9, 1).Value = "Launch On Cut"
    Cells(8, 1).Value = 0
End Sub

Private Sub RefreshList_Click()
    Call RefreshWireListSettings
End Sub

Private Sub Reset_Click()
    Dim r1 As Range, cell As Range, a As Variant
    Set r1 = Range(Cells(4, 2), Cells(21, 5))
        For Each cell In r1
           cell.Value = ""
        Next cell
        With Cells(2, 1)
        .Value = "(Insert wire name here)"
        .Font.Size = 20
        End With
        ComboBox1.Value = ""
End Sub

Private Sub OptionButton1_Click()
   Cells(4, 1).Value = "standard"
   Cells(5, 1).Value = 0
   Cells(6, 1).Value = 0
End Sub

Private Sub OptionButton2_Click()
    Cells(5, 1).Value = "critical"
    Cells(4, 1).Value = 0
    Cells(6, 1).Value = 0
End Sub

Private Sub combobox1Add_()
    ComboBox1.ListRows = ComboBox1.ListRows + 1
    ComboBox1.AddItem ws1.Cells(2, 1).Value
End Sub



Private Sub Add_Click()
    Call AddWire
    Dim a As Integer, alreadythere As Boolean
    For a = 0 To ComboBox1.ListCount + -1
        If Cells(2, 1).Value = ComboBox1.list(a) Then alreadythere = True
    Next a
      If alreadythere = False Then
        ComboBox1.ListRows = ComboBox1.ListRows + 1
        ComboBox1.AddItem Cells(2, 1).Value
    End If

End Sub

Private Sub Edit_Click()
    Call editwire
End Sub


Sub Clear()
Dim a As Integer
     OLEObjects("combobox1").object.RemoveItem 0

End Sub



