VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Call AddToCellMenu

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call RegisterChange(Target)

End Sub

Private Sub Worksheet_Deactivate()
    Call DeleteFromCellMenu

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Count > 10000 Then Exit Sub
    LastValueSelected = Application.Transpose(Target.value)
    LastCommentsSelected = GetComments(Target)

End Sub
