Attribute VB_Name = "Z1_Parameters_ContextMenu"
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Restore last custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
        .OnAction = "RevertToLastValue"
        .FaceId = 155
        .Caption = "Restaurer la valeur précédente"
        .Tag = "EDIT_Control_Tag"
    End With
    
    ' Restore first custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "RestoreFirstValue"
        .FaceId = 154
        .Caption = "Restaurer la valeur d'origine"
        .Tag = "EDIT_Control_Tag"
    End With
   
    
    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=3
    ' Add a separator to the Cell context menu.
    
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "EDIT_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

