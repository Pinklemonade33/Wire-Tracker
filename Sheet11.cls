VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Deactivate()
Me.Visible = xlSheetHidden
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
