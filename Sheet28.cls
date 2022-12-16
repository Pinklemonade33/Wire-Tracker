VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim dropexit As Boolean

Sub test()
ActiveSheet.OLEObjects("hSettingsLabel").object.ForeColor = RGB(242, 242, 242)
ActiveSheet.OLEObjects("hSettingsLabel").object.BackColor = RGB(0, 32, 96)
End Sub

Private Sub hAdd_Click()
AddWire.Height = 370.2
AddWire.Width = 141
AddWire.Show
End Sub

Private Sub hCalculation_Click()
ThisWorkbook.Worksheets("Calculate").Activate
End Sub

Private Sub hRemove_Click()
RemoveWire.Height = 421.2
RemoveWire.Width = 403.8
RemoveWire.rHighCuts = True
RemoveWire.rBulk = True
RemoveWire.rLowCuts = True
RemoveWire.Show
End Sub

Private Sub hSettings_Click()
Settings.Show
End Sub

Private Sub hView_Click()
View.Width = 627.6
View.Height = 489.6
View.vcHighCuts = True
View.vcLowCuts = True
View.vcBulk = True
View.vfHighCuts = True
View.vfLowCuts = True
View.vfBulk = True
Call FillAltCombo
View.Show
End Sub

Sub FillAltCombo()
View.vfAltCombo.object.AddItem "Pending"
View.vfAltCombo.object.AddItem "Picked"
View.vfAltCombo.object.AddItem "Removed"
View.vfAltCombo.object.AddItem "Added"
End Sub
