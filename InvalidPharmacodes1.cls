VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InvalidPharmacodes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If LastEditedCell Is Nothing Then LastEditedCell = ActiveCell
    Call ColorLabelling(LastEditedCell)

End Sub

Private Sub Worksheet_Deactivate()
    Set LastEditedCell = Nothing

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Set LastEditedCell = ActiveCell

End Sub